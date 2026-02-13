const { app, BrowserWindow, screen, ipcMain } = require('electron');
const path = require('node:path');
const { execFile } = require('node:child_process');

const iconPath = path.join(__dirname, '../assets', process.platform === 'win32' ? 'app-icon.ico' : 'app-icon.png');
const { registerHandlers } = require('./ipc-handlers');
const { installExcelAddin } = require('./excel-addin-installer');

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
let excelWindowMonitor = null;
let topmostReassertTimers = [];
let currentExcelOwnerHwnd = null;

/**
 * Get the main window reference (used by ipc-handlers)
 * @returns {BrowserWindow|null}
 */
function getMainWindow() {
  return mainWindow;
}

function clearTopmostReassertTimers() {
  topmostReassertTimers.forEach((timer) => clearTimeout(timer));
  topmostReassertTimers = [];
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

function setWindowOwnerWin32(win, ownerHwnd) {
  if (process.platform !== 'win32') {
    return;
  }
  const hwnd = getNativeWindowHandleValue(win);
  if (hwnd === null || hwnd === undefined) {
    return;
  }

  const hwndValue = hwnd.toString();
  const ownerValue = ownerHwnd ? ownerHwnd.toString() : '0';
  const script = [
    '$ErrorActionPreference = "SilentlyContinue";',
    'Add-Type -TypeDefinition @\'',
    'using System;',
    'using System.Runtime.InteropServices;',
    'public class User32 {',
    '  [DllImport("user32.dll", SetLastError=true)] public static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);',
    '  [DllImport("user32.dll", SetLastError=true)] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);',
    '}',
    '\'@ | Out-Null;',
    '$GWL_HWNDPARENT = -8;',
    '$SWP_NOMOVE = 0x0002;',
    '$SWP_NOSIZE = 0x0001;',
    '$SWP_NOACTIVATE = 0x0010;',
    '$SWP_SHOWWINDOW = 0x0040;',
    `$hWnd = [IntPtr]::new(${hwndValue});`,
    `$owner = [IntPtr]::new(${ownerValue});`,
    '[User32]::SetWindowLongPtr($hWnd, $GWL_HWNDPARENT, $owner) | Out-Null;',
    '[User32]::SetWindowPos($hWnd, [IntPtr]::Zero, 0, 0, 0, 0, ' +
      '$SWP_NOMOVE -bor $SWP_NOSIZE -bor $SWP_NOACTIVATE -bor $SWP_SHOWWINDOW) | Out-Null;'
  ].join('\n');

  execFile(
    'powershell.exe',
    [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-EncodedCommand',
      Buffer.from(script, 'utf16le').toString('base64')
    ],
    { windowsHide: true, timeout: 1500 },
    () => {}
  );
}

function enforceTopmostForExcel(win) {
  if (!win || win.isDestroyed()) {
    return;
  }
  win.setAlwaysOnTop(true, 'screen-saver');
  try {
    win.moveTop();
  } catch (error) {
    // Ignore z-order errors.
  }
  try {
    win.showInactive();
  } catch (error) {
    // Ignore visibility errors.
  }
}

function applyExcelForegroundState(win, state) {
  if (!win || win.isDestroyed()) {
    return;
  }

  const isExcelActive = Boolean(state && state.isExcelActive);
  const excelHwnd = state && state.hwnd ? state.hwnd : null;

  clearTopmostReassertTimers();

  if (!isExcelActive) {
    if (currentExcelOwnerHwnd !== null) {
      setWindowOwnerWin32(win, null);
      currentExcelOwnerHwnd = null;
    }
    win.setAlwaysOnTop(false);
    return;
  }

  if (excelHwnd && currentExcelOwnerHwnd !== excelHwnd) {
    setWindowOwnerWin32(win, excelHwnd);
    currentExcelOwnerHwnd = excelHwnd;
  } else if (!excelHwnd && currentExcelOwnerHwnd !== null) {
    setWindowOwnerWin32(win, null);
    currentExcelOwnerHwnd = null;
  }

  // Immediate assert reduces visible flicker during Excel activation.
  enforceTopmostForExcel(win);

  // Short, finite reassert burst to win activation races without constant churn.
  [35, 90, 170, 280, 420].forEach((delayMs) => {
    const timer = setTimeout(() => enforceTopmostForExcel(win), delayMs);
    if (typeof timer.unref === 'function') {
      timer.unref();
    }
    topmostReassertTimers.push(timer);
  });
}

function startExcelWindowMonitor(win) {
  if (process.platform !== 'win32') {
    return;
  }
  if (!win || win.isDestroyed()) {
    return;
  }
  if (excelWindowMonitor) {
    return;
  }

  try {
    // WinEvent foreground hook monitor (no polling).
    const ExcelForegroundHook = require('./excel-foreground-hook');
    excelWindowMonitor = new ExcelForegroundHook((state) => {
      applyExcelForegroundState(win, state);
    });
    excelWindowMonitor.start();
  } catch (error) {
    excelWindowMonitor = null;
    console.error(`[WindowMonitor] Foreground hook unavailable: ${error.message}`);
  }
}

function stopExcelWindowMonitor() {
  if (excelWindowMonitor) {
    try {
      excelWindowMonitor.stop();
    } catch (error) {
      console.error(`[WindowMonitor] Failed to stop cleanly: ${error.message}`);
    } finally {
      excelWindowMonitor = null;
    }
  }

  clearTopmostReassertTimers();
  if (mainWindow && !mainWindow.isDestroyed()) {
    if (currentExcelOwnerHwnd !== null) {
      setWindowOwnerWin32(mainWindow, null);
      currentExcelOwnerHwnd = null;
    }
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
      console.log(`[Window] alwaysOnTop: ${wasOnTop} -> ${value}`);
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
  const { width, height, x, y } = primaryDisplay.workArea;

  // Window dimensions
  const WIN_WIDTH = 620;
  const WIN_HEIGHT = 580;
  const RIGHT_MARGIN = 50;
  const BOTTOM_MARGIN = 150;

  mainWindow = new BrowserWindow({
    width: WIN_WIDTH,
    height: WIN_HEIGHT,
    x: x + width - WIN_WIDTH - RIGHT_MARGIN,
    y: y + height - WIN_HEIGHT - BOTTOM_MARGIN,
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

  // Load from Vite dev server in development, built files in production
  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
    // Open DevTools in development
    mainWindow.webContents.openDevTools();
  } else {
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
  }

  // Clean up reference when window is closed
  mainWindow.on('closed', () => {
    stopExcelWindowMonitor();
    mainWindow = null;
  });

  startExcelWindowMonitor(mainWindow);
}

module.exports = { getMainWindow };
