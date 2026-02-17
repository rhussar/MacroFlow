const { app, BrowserWindow, screen, ipcMain } = require('electron');
const path = require('node:path');
const { execFile } = require('node:child_process');
const { promisify } = require('node:util');

const iconPath = path.join(__dirname, '../assets', process.platform === 'win32' ? 'app-icon.ico' : 'app-icon.png');
const { registerHandlers } = require('./ipc-handlers');
const { installExcelAddin } = require('./excel-addin-installer');
const logger = require('./logger');

const execFileAsync = promisify(execFile);

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
let lastExcelDisplayId = null;
let hasInitialContextPlacement = false;

async function writeRegistryStringValue(keyPath, valueName, valueData) {
  const args = ['add', keyPath];
  if (valueName === null) {
    args.push('/ve');
  } else {
    args.push('/v', valueName);
  }
  args.push('/t', 'REG_SZ', '/d', valueData, '/f');

  await execFileAsync('reg.exe', args, { windowsHide: true });
}

async function registerInstallPathInRegistry() {
  if (process.platform !== 'win32' || !app.isPackaged) {
    return;
  }

  const exePath = process.execPath;
  const appPathsKey = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\MacroFlow.exe';

  try {
    await writeRegistryStringValue('HKCU\\Software\\MacroFlow', 'InstallPath', exePath);
    await writeRegistryStringValue(appPathsKey, null, exePath);
    await writeRegistryStringValue(appPathsKey, 'Path', path.dirname(exePath));

    logger.info('[InstallPath] registry updated', { exePath });
  } catch (error) {
    logger.warn('[InstallPath] registry update failed', { exePath, error: error.message });
  }
}

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

function getDisplayById(id) {
  if (id === null || id === undefined) {
    return null;
  }

  return screen.getAllDisplays().find((display) => display.id === id) || null;
}

function getCursorDisplay() {
  try {
    const cursorPoint = screen.getCursorScreenPoint();
    return screen.getDisplayNearestPoint(cursorPoint);
  } catch {
    return null;
  }
}

function getDisplayFromRect(rect) {
  if (!rect || typeof rect !== 'object') {
    return null;
  }

  const left = Number(rect.left);
  const top = Number(rect.top);
  const right = Number(rect.right);
  const bottom = Number(rect.bottom);

  if ([left, top, right, bottom].some((value) => Number.isNaN(value))) {
    return null;
  }

  const center = {
    x: Math.round((left + right) / 2),
    y: Math.round((top + bottom) / 2)
  };

  return screen.getDisplayNearestPoint(center);
}

function resolveTargetDisplay(win, options = {}) {
  if (options.preferExcel && lastExcelDisplayId !== null) {
    const excelDisplay = getDisplayById(lastExcelDisplayId);
    if (excelDisplay) {
      return excelDisplay;
    }
    lastExcelDisplayId = null;
  }

  if (options.preferCursor) {
    const cursorDisplay = getCursorDisplay();
    if (cursorDisplay) {
      return cursorDisplay;
    }
  }

  if (win && !win.isDestroyed()) {
    return screen.getDisplayMatching(win.getBounds());
  }

  return screen.getPrimaryDisplay();
}

function applyAdaptiveLayout(win, reason = 'unspecified', options = {}) {
  if (!win || win.isDestroyed()) {
    return;
  }

  const display = resolveTargetDisplay(win, options);
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
    height: layout.height,
    preferExcel: Boolean(options.preferExcel),
    preferCursor: Boolean(options.preferCursor)
  });
}

function scheduleAdaptiveLayout(win, reason, delayMs = 100, options = {}) {
  if (adaptiveLayoutTimer) {
    clearTimeout(adaptiveLayoutTimer);
    adaptiveLayoutTimer = null;
  }

  adaptiveLayoutTimer = setTimeout(() => {
    adaptiveLayoutTimer = null;
    applyAdaptiveLayout(win, reason, options);
  }, delayMs);

  if (typeof adaptiveLayoutTimer.unref === 'function') {
    adaptiveLayoutTimer.unref();
  }
}

function attachAdaptiveWindowListeners(win) {
  const onDisplayMetricsChanged = () => scheduleAdaptiveLayout(win, 'display-metrics-changed', 120, { preferExcel: true });
  const onDisplayAdded = () => scheduleAdaptiveLayout(win, 'display-added', 120, { preferExcel: true });
  const onDisplayRemoved = () => scheduleAdaptiveLayout(win, 'display-removed', 120, { preferExcel: true });
  const onWindowShow = () => scheduleAdaptiveLayout(win, 'window-show', 80, { preferExcel: true, preferCursor: true });
  const onWindowRestore = () => scheduleAdaptiveLayout(win, 'window-restore', 80, { preferExcel: true, preferCursor: true });

  screen.on('display-metrics-changed', onDisplayMetricsChanged);
  screen.on('display-added', onDisplayAdded);
  screen.on('display-removed', onDisplayRemoved);
  win.on('show', onWindowShow);
  win.on('restore', onWindowRestore);

  detachDisplayListeners = () => {
    screen.removeListener('display-metrics-changed', onDisplayMetricsChanged);
    screen.removeListener('display-added', onDisplayAdded);
    screen.removeListener('display-removed', onDisplayRemoved);
    if (!win.isDestroyed()) {
      win.removeListener('show', onWindowShow);
      win.removeListener('restore', onWindowRestore);
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
  mainWindow.__macroflowHelperManagedTopmost = false;
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
  mainWindow.__macroflowHelperManagedTopmost = false;
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

  // Token-based reassert mechanism: delayed setAlwaysOnTop(true) calls
  // through Electron's API to ensure topmost sticks after window transitions.
  // Direct Win32 SetWindowPos calls from the C# helper can be overridden by
  // Electron's internal window management, so we mirror reasserts here.
  let jsReassertToken = 0;
  const JS_REASSERT_DELAYS = [50, 150, 350];
  let helperTargetConfirmed = false;

  try {
    const hwnd = getNativeWindowHandleValue(win);
    if (hwnd === null) {
      throw new Error('Main window handle unavailable.');
    }

    const targetHwnd = hwnd.toString();
    const WindowFocusHelperClient = require('./window-focus-helper-client');
    windowFocusHelper = new WindowFocusHelperClient({
      logger,
      maxRestartAttempts: 3,
      onStateChange: (state) => {
        const stateTargetHwnd = state && state.targetHwnd ? String(state.targetHwnd) : null;
        if (stateTargetHwnd !== targetHwnd) {
          if (helperTargetConfirmed) {
            helperTargetConfirmed = false;
          }
          logger.debug('[WindowHelper] ignoring mismatched helper state', {
            expectedTargetHwnd: targetHwnd,
            stateTargetHwnd
          });
          return;
        }

        if (!helperTargetConfirmed) {
          helperTargetConfirmed = true;
          logger.debug('[WindowHelper] helper target confirmed', { targetHwnd });
        }

        logger.debug('[WindowHelper] state', state);

        if (helperFallbackMode) {
          disableFallbackAlwaysOnTop('helper-state-received');
        }

        // Increment token on every state change to cancel stale reasserts
        jsReassertToken += 1;
        const currentToken = jsReassertToken;

        if (win && !win.isDestroyed()) {
          const excelActive = Boolean(state && state.excelActive);
          const stateProcess = String((state && state.process) || '').toLowerCase();

          const isSelfProcess = ['electron', 'macroflow'].includes(stateProcess);
          const branch = excelActive
            ? 'excel-active'
            : (isSelfProcess ? 'electron-transient-ignored' : 'nonexcel-decisive');

          if (branch === 'excel-active') {
            logger.debug('[WindowHelper] excel-active', { process: stateProcess, token: currentToken });
            win.setAlwaysOnTop(true);

            // When Excel becomes active, schedule delayed reasserts through
            // Electron's API. This ensures topmost sticks even if Electron
            // overrides the C# helper's direct Win32 SetWindowPos calls.
            for (const delay of JS_REASSERT_DELAYS) {
              setTimeout(() => {
                if (currentToken !== jsReassertToken) return;
                if (!win || win.isDestroyed()) return;
                win.setAlwaysOnTop(true);
                logger.debug('[WindowHelper] JS reassert applied', { delay, token: currentToken });
              }, delay);
            }
          } else if (branch === 'nonexcel-decisive') {
            logger.debug('[WindowHelper] nonexcel-decisive', { process: stateProcess, token: currentToken });
            win.setAlwaysOnTop(false);
            // No demote reasserts — C# helper owns Z-order placement via
            // DemoteWindow. Repeated setAlwaysOnTop(false) calls fight with
            // the helper's SetWindowPos(target, fg) by placing the window
            // back at the top of all non-topmost windows.
          } else {
            logger.debug('[WindowHelper] electron-transient-ignored', {
              process: stateProcess,
              token: currentToken
            });
          }
        }

        if (state && state.excelActive && state.excelRect) {
          const excelDisplay = getDisplayFromRect(state.excelRect);
          if (excelDisplay) {
            lastExcelDisplayId = excelDisplay.id;

            if (!hasInitialContextPlacement) {
              scheduleAdaptiveLayout(win, 'excel-display-anchor', 80, { preferExcel: true });
              hasInitialContextPlacement = true;
            }
          }
        } else if (!hasInitialContextPlacement) {
          scheduleAdaptiveLayout(win, 'initial-cursor-anchor', 80, { preferCursor: true });
          hasInitialContextPlacement = true;
        }
      },
      onError: (message, payload) => {
        if (payload && payload.type === 'process-exit') {
          helperTargetConfirmed = false;
        }
        if (payload && typeof payload === 'object') {
          logger.warn('[WindowHelper] warning', payload);
          return;
        }
        logger.warn('[WindowHelper] warning', { message });
      },
      onFatal: (message) => {
        helperTargetConfirmed = false;
        logger.error('[WindowHelper] fatal', { message });
        if (windowFocusHelper) {
          try {
            windowFocusHelper.stop();
          } catch {
            // Ignore shutdown errors in fatal path.
          }
          windowFocusHelper = null;
        }
        if (win && !win.isDestroyed()) {
          win.__macroflowHelperManagedTopmost = false;
        }
        enableFallbackAlwaysOnTop(message);
      }
    });

    windowFocusHelper.start(targetHwnd);
    if (windowFocusHelper) {
      win.__macroflowHelperManagedTopmost = true;
      logger.info('[WindowMonitor] helper monitor started', { hwnd: targetHwnd });
    }
  } catch (error) {
    windowFocusHelper = null;
    win.__macroflowHelperManagedTopmost = false;
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

  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.__macroflowHelperManagedTopmost = false;
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
    // Register window control handlers BEFORE creating window
    registerWindowHandlers();

    registerHandlers();
    createWindow();

    // Persist current installed exe path so VBA launcher is location-independent.
    registerInstallPathInRegistry();

    // Install/update Excel add-in in the background so first paint is fast.
    installExcelAddin().catch((error) => {
      logger.error('[AddinInstaller] background install failed', { error: error.message });
    });

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
  const launchDisplay = resolveTargetDisplay(null, { preferExcel: true, preferCursor: true });
  const layout = getAdaptiveLayout(launchDisplay);
  const { workArea } = launchDisplay;

  const initialX = workArea.x + Math.max(0, workArea.width - layout.width - layout.rightMargin);
  const initialY = workArea.y + Math.max(0, workArea.height - layout.height - layout.bottomMargin);

  mainWindow = new BrowserWindow({
    width: layout.width,
    height: layout.height,
    x: initialX,
    y: initialY,
    minWidth: WINDOW_BASELINE.minWidth,
    minHeight: WINDOW_BASELINE.minHeight,
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

  mainWindow.__macroflowHelperManagedTopmost = false;
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
    applyAdaptiveLayout(mainWindow, 'did-finish-load', { preferExcel: true, preferCursor: true });
    // Start the helper AFTER the window is fully loaded so that
    // setAlwaysOnTop toggles operate on a fully-ready native window.
    // Starting earlier causes Electron's internal alwaysOnTop state to
    // initialise incorrectly, making the first demotion/promotion cycle fail.
    startExcelWindowMonitor(mainWindow);
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
    hasInitialContextPlacement = false;
    lastExcelDisplayId = null;
  });
}

module.exports = { getMainWindow };






