const { app, BrowserWindow, screen, ipcMain } = require('electron');
const path = require('node:path');

const iconPath = path.join(__dirname, '../assets', process.platform === 'win32' ? 'app-icon.ico' : 'app-icon.png');
const { registerHandlers } = require('./ipc-handlers');
const { installExcelAddin } = require('./excel-addin-installer');
const logger = require('./logger');

const WINDOW_BASELINE = {
  displayWidth: 1920,
  displayHeight: 1080,
  width: 620,
  height: 580,
  rightMargin: 50,
  bottomMargin: 150,
  minScale: 0.6,
  maxScale: 1,
  minWidth: 420,
  minHeight: 390,
  minRightMargin: 16,
  minBottomMargin: 20
};

// CRITICAL: Handle Squirrel installer events FIRST (must be before any other code)
if (require('electron-squirrel-startup')) {
  app.quit();
  return;
}

// Ensure Windows uses our app identity for taskbar grouping (helps avoid a stale pinned/shortcut icon).
if (process.platform === 'win32') {
  app.setAppUserModelId('com.macroflow.desktop');
}

// Determine if running in development mode
const isDev = process.env.NODE_ENV === 'development';

// Store reference to main window for IPC handlers
let mainWindow = null;
let windowFocusHelper = null;
let helperFallbackMode = false;
let adaptiveLayoutTimer = null;
let lastAdaptiveLayoutKey = '';
let detachDisplayListeners = null;

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function getDisplayScale(display) {
  const widthFactor = display.workArea.width / WINDOW_BASELINE.displayWidth;
  const heightFactor = display.workArea.height / WINDOW_BASELINE.displayHeight;
  const rawScale = Math.min(widthFactor, heightFactor);
  return clamp(rawScale, WINDOW_BASELINE.minScale, WINDOW_BASELINE.maxScale);
}

function getAdaptiveLayout(display) {
  const scale = getDisplayScale(display);
  return {
    scale,
    width: Math.max(WINDOW_BASELINE.minWidth, Math.round(WINDOW_BASELINE.width * scale)),
    height: Math.max(WINDOW_BASELINE.minHeight, Math.round(WINDOW_BASELINE.height * scale)),
    rightMargin: Math.max(WINDOW_BASELINE.minRightMargin, Math.round(WINDOW_BASELINE.rightMargin * scale)),
    bottomMargin: Math.max(WINDOW_BASELINE.minBottomMargin, Math.round(WINDOW_BASELINE.bottomMargin * scale))
  };
}

function getDisplayForWindow(win) {
  const bounds = win.getBounds();
  return screen.getDisplayMatching(bounds);
}

function applyAdaptiveLayout(win, reason = 'unspecified') {
  if (!win || win.isDestroyed()) {
    return;
  }

  const display = getDisplayForWindow(win);
  const workArea = display.workArea;
  const layout = getAdaptiveLayout(display);

  const x = workArea.x + Math.max(0, workArea.width - layout.width - layout.rightMargin);
  const y = workArea.y + Math.max(0, workArea.height - layout.height - layout.bottomMargin);
  const layoutKey = `${display.id}:${layout.scale}:${layout.width}:${layout.height}:${x}:${y}`;

  if (layoutKey !== lastAdaptiveLayoutKey) {
    win.setBounds({ x, y, width: layout.width, height: layout.height });
    lastAdaptiveLayoutKey = layoutKey;
  }

  if (win.webContents && !win.webContents.isDestroyed()) {
    win.webContents.setZoomFactor(layout.scale);
  }

  logger.debug('[WindowScale] adaptive layout applied', {
    reason,
    displayId: display.id,
    scale: layout.scale,
    width: layout.width,
    height: layout.height
  });
}

function scheduleAdaptiveLayout(win, reason, delayMs = 100) {
  if (adaptiveLayoutTimer) {
    clearTimeout(adaptiveLayoutTimer);
    adaptiveLayoutTimer = null;
  }

  adaptiveLayoutTimer = setTimeout(() => {
    adaptiveLayoutTimer = null;
    applyAdaptiveLayout(win, reason);
  }, delayMs);

  if (typeof adaptiveLayoutTimer.unref === 'function') {
    adaptiveLayoutTimer.unref();
  }
}

function attachAdaptiveWindowListeners(win) {
  const onDisplayMetricsChanged = () => scheduleAdaptiveLayout(win, 'display-metrics-changed', 120);
  const onDisplayAdded = () => scheduleAdaptiveLayout(win, 'display-added', 120);
  const onDisplayRemoved = () => scheduleAdaptiveLayout(win, 'display-removed', 120);
  const onWindowMoved = () => scheduleAdaptiveLayout(win, 'window-moved', 120);

  screen.on('display-metrics-changed', onDisplayMetricsChanged);
  screen.on('display-added', onDisplayAdded);
  screen.on('display-removed', onDisplayRemoved);
  win.on('move', onWindowMoved);

  detachDisplayListeners = () => {
    screen.removeListener('display-metrics-changed', onDisplayMetricsChanged);
    screen.removeListener('display-added', onDisplayAdded);
    screen.removeListener('display-removed', onDisplayRemoved);
    if (!win.isDestroyed()) {
      win.removeListener('move', onWindowMoved);
    }
  };
}

/**
 * Get the main window reference (used by ipc-handlers)
 * @returns {BrowserWindow|null}
 */
function getMainWindow() {
  return mainWindow;
}

function getNativeWindowHandleValue(win) {
  if (!win || win.isDestroyed()) {
    return null;
  }
  const handle = win.getNativeWindowHandle();
  if (!handle || handle.length === 0) {
    return null;
  }
  if (handle.length === 8) {
    return handle.readBigUInt64LE(0);
  }
  if (handle.length === 4) {
    return BigInt(handle.readUInt32LE(0));
  }
  return null;
}

function enableFallbackAlwaysOnTop(reason) {
  if (!mainWindow || mainWindow.isDestroyed()) {
    return;
  }

  helperFallbackMode = true;
  mainWindow.setAlwaysOnTop(true);
  logger.warn('[WindowMonitor] fallback always-on-top enabled', { reason });
}

function disableFallbackAlwaysOnTop(reason) {
  if (!mainWindow || mainWindow.isDestroyed()) {
    helperFallbackMode = false;
    return;
  }

  if (!helperFallbackMode) {
    return;
  }

  helperFallbackMode = false;
  mainWindow.setAlwaysOnTop(false);
  logger.info('[WindowMonitor] fallback always-on-top disabled', { reason });
}

function startExcelWindowMonitor(win) {
  if (process.platform !== 'win32') {
    return;
  }
  if (!win || win.isDestroyed()) {
    return;
  }
  if (windowFocusHelper) {
    return;
  }

  try {
    const hwnd = getNativeWindowHandleValue(win);
    if (hwnd === null) {
      throw new Error('Main window handle unavailable.');
    }

    const WindowFocusHelperClient = require('./window-focus-helper-client');
    windowFocusHelper = new WindowFocusHelperClient({
      logger,
      maxRestartAttempts: 3,
      onStateChange: (state) => {
        logger.debug('[WindowHelper] state', state);
        if (helperFallbackMode) {
          disableFallbackAlwaysOnTop('helper-state-received');
        }
      },
      onError: (message) => {
        logger.warn('[WindowHelper] warning', { message });
      },
      onFatal: (message) => {
        logger.error('[WindowHelper] fatal', { message });
        if (windowFocusHelper) {
          try {
            windowFocusHelper.stop();
          } catch {
            // Ignore shutdown errors in fatal path.
          }
          windowFocusHelper = null;
        }
        enableFallbackAlwaysOnTop(message);
      }
    });

    windowFocusHelper.start(hwnd.toString());
    if (windowFocusHelper) {
      logger.info('[WindowMonitor] helper monitor started', { hwnd: hwnd.toString() });
    }
  } catch (error) {
    windowFocusHelper = null;
    logger.error('[WindowMonitor] failed to start helper monitor', { error: error.message });
    enableFallbackAlwaysOnTop(error.message);
  }
}

function stopExcelWindowMonitor() {
  if (windowFocusHelper) {
    try {
      windowFocusHelper.stop();
    } catch (error) {
      logger.error('[WindowMonitor] failed to stop helper monitor', { error: error.message });
    } finally {
      windowFocusHelper = null;
    }
  }

  disableFallbackAlwaysOnTop('monitor-stop');

  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.setAlwaysOnTop(false);
  }
}

// CRITICAL: Request single instance lock - prevent multiple windows
const gotLock = app.requestSingleInstanceLock();

if (!gotLock) {
  // Another instance is already running - quit immediately without doing anything
  app.quit();
  // IMPORTANT: Do not execute any code after this
} else {
  // We have the lock - set up the single instance behavior

  // Handle second-instance attempts by focusing the existing window
  app.on('second-instance', () => {
    const windows = BrowserWindow.getAllWindows();
    if (windows.length > 0) {
      const win = windows[0];
      if (win.isMinimized()) {
        win.restore();
      }
      win.focus();
    }
  });

  // ONLY register app lifecycle events if we have the lock
  app.whenReady().then(async () => {
    // Install Excel Add-in before anything else
    await installExcelAddin();

    // Register window control handlers BEFORE creating window
    registerWindowHandlers();

    registerHandlers();
    createWindow();

    // On macOS, re-create window when dock icon is clicked and no windows are open
    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
      }
    });
  });

  // Quit when all windows are closed (except on macOS)
  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
      app.quit();
    }
  });

  app.on('before-quit', () => {
    stopExcelWindowMonitor();
  });
}

/**
 * Register IPC handlers for window control (alwaysOnTop, etc.)
 * These are registered separately to avoid circular dependencies
 */
function registerWindowHandlers() {
  // Set alwaysOnTop state
  ipcMain.handle('window:setAlwaysOnTop', (_, value) => {
    const win = getMainWindow();
    if (win && !win.isDestroyed()) {
      const wasOnTop = win.isAlwaysOnTop();
      win.setAlwaysOnTop(Boolean(value));
      logger.info('[Window] alwaysOnTop changed', {
        previousValue: wasOnTop,
        currentValue: Boolean(value)
      });
      return { success: true, previousValue: wasOnTop, currentValue: Boolean(value) };
    }
    return { success: false, message: 'Window not available' };
  });

  // Get alwaysOnTop state
  ipcMain.handle('window:getAlwaysOnTop', () => {
    const win = getMainWindow();
    if (win && !win.isDestroyed()) {
      return { success: true, value: win.isAlwaysOnTop() };
    }
    return { success: false, message: 'Window not available' };
  });
}

// Window creation function
function createWindow() {
  const primaryDisplay = screen.getPrimaryDisplay();
  const layout = getAdaptiveLayout(primaryDisplay);
  const { workArea } = primaryDisplay;

  const initialX = workArea.x + Math.max(0, workArea.width - layout.width - layout.rightMargin);
  const initialY = workArea.y + Math.max(0, workArea.height - layout.height - layout.bottomMargin);

  mainWindow = new BrowserWindow({
    width: layout.width,
    height: layout.height,
    x: initialX,
    y: initialY,
    frame: false,
    transparent: false,
    hasShadow: false,
    roundedCorners: false,
    alwaysOnTop: false,
    resizable: true,
    movable: true,
    skipTaskbar: false,
    backgroundColor: '#1e1e1e',
    title: 'MacroFlow',
    icon: iconPath,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      sandbox: false
    }
  });

  mainWindow.webContents.setZoomFactor(layout.scale);
  attachAdaptiveWindowListeners(mainWindow);

  // Load from Vite dev server in development, built files in production
  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
    // Open DevTools in development
    mainWindow.webContents.openDevTools();
  } else {
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
  }

  mainWindow.webContents.on('did-finish-load', () => {
    applyAdaptiveLayout(mainWindow, 'did-finish-load');
  });

  // Clean up reference when window is closed
  mainWindow.on('closed', () => {
    if (detachDisplayListeners) {
      detachDisplayListeners();
      detachDisplayListeners = null;
    }
    if (adaptiveLayoutTimer) {
      clearTimeout(adaptiveLayoutTimer);
      adaptiveLayoutTimer = null;
    }
    stopExcelWindowMonitor();
    mainWindow = null;
    lastAdaptiveLayoutKey = '';
  });

  startExcelWindowMonitor(mainWindow);
}

module.exports = { getMainWindow };


