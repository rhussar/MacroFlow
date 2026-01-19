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
  // Get screen dimensions for phantom taskpane positioning
  const primaryDisplay = screen.getPrimaryDisplay();
  const { width, height, x, y } = primaryDisplay.workArea;

  // Safe Zone gaps to keep Excel controls accessible
  const TOP_GAP = 240;    // Leave space for Excel title bar, ribbon, tabs, column headers
  const BOTTOM_GAP = 50;  // Leave space for Excel status bar and zoom slider
  const SIDEBAR_WIDTH = 400;

  const mainWindow = new BrowserWindow({
    width: SIDEBAR_WIDTH,
    height: height - TOP_GAP - BOTTOM_GAP, // Reduce height for safe zones
    x: x + width - SIDEBAR_WIDTH, // Snap to far right
    y: y + TOP_GAP, // Push down from top
    frame: false, // Remove Windows title bar
    alwaysOnTop: true, // Float above Excel
    resizable: false, // Fixed width like real taskpane
    movable: false, // Lock in place
    skipTaskbar: false, // Keep visible in taskbar for now
    title: 'MacroFlow',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      sandbox: false // Required for COM
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
