const { app, BrowserWindow, screen } = require('electron');
const path = require('node:path');
const { registerHandlers } = require('./ipc-handlers');
const { installExcelAddin } = require('./excel-addin-installer');

// CRITICAL: Handle Squirrel installer events FIRST (must be before any other code)
if (require('electron-squirrel-startup')) {
  app.quit();
  return;
}

// Determine if running in development mode
const isDev = process.env.NODE_ENV === 'development';

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
      const mainWindow = windows[0];
      if (mainWindow.isMinimized()) {
        mainWindow.restore();
      }
      mainWindow.focus();
    }
  });

  // ONLY register app lifecycle events if we have the lock
  app.whenReady().then(async () => {
    // Install Excel Add-in before anything else
    await installExcelAddin();

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
}

// Window creation function
function createWindow() {
  const primaryDisplay = screen.getPrimaryDisplay();
  const { width, height, x, y } = primaryDisplay.workArea;

  // Window dimensions
  const WIN_WIDTH = 750;
  const WIN_HEIGHT = 738;
  const RIGHT_MARGIN = 21;
  const BOTTOM_MARGIN = 55;

  const mainWindow = new BrowserWindow({
    width: WIN_WIDTH,
    height: WIN_HEIGHT,
    x: x + width - WIN_WIDTH - RIGHT_MARGIN,
    y: y + height - WIN_HEIGHT - BOTTOM_MARGIN,
    frame: false,
    alwaysOnTop: true,
    resizable: false,
    movable: false,
    skipTaskbar: false,
    title: 'MacroFlow',
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
}
