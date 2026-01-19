console.log('[DEBUG] Loading electron...')
const electron = require('electron')
console.log('[DEBUG] Electron type:', typeof electron)
console.log('[DEBUG] Electron value:', electron)
console.log('[DEBUG] Electron.app:', electron.app)

const { app, BrowserWindow, ipcMain } = electron
const path = require('path')

console.log('[DEBUG] app:', app)
console.log('[DEBUG] BrowserWindow:', typeof BrowserWindow)

/**
 * Test Excel COM connection
 */
async function testExcelConnection() {
  try {
    const winax = require('winax')
    const excel = new winax.Object('Excel.Application', {
      activate: true,
      type: true
    })
    excel.Visible = true
    const workbook = excel.Workbooks.Add()
    const worksheet = workbook.ActiveSheet
    const cellA1 = worksheet.Range('A1')
    cellA1.Value = 'Hello World'
    cellA1.Font.Bold = true
    cellA1.Font.Size = 14
    cellA1.Font.Color = 0x0000FF // Red in BGR format
    return 'Successfully wrote "Hello World" to Excel cell A1!'
  } catch (error) {
    if (error instanceof Error) {
      if (error.message.includes('winax')) {
        throw new Error('COM library not available. Make sure winax is installed: npm install winax')
      }
      throw new Error(`Excel COM Error: ${error.message}`)
    }
    throw error
  }
}

let mainWindow = null

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 600,
    height: 400,
    title: 'MacroFlow - Proof of Life',
    webPreferences: {
      preload: path.join(__dirname, '../preload/index.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false // Required for COM
    }
  })

  // Load the app
  if (process.env.ELECTRON_RENDERER_URL) {
    mainWindow.loadURL(process.env.ELECTRON_RENDERER_URL)
  } else {
    mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'))
  }
}

// App lifecycle
app.whenReady().then(() => {
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

// IPC Handler for Excel COM test
ipcMain.handle('test-excel-connection', async () => {
  try {
    const result = await testExcelConnection()
    return { success: true, message: result }
  } catch (error) {
    return {
      success: false,
      message: error instanceof Error ? error.message : 'Unknown error'
    }
  }
})
