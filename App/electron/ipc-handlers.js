const { ipcMain } = require('electron');
const excel = require('./excel-service');

function registerHandlers() {
  // Test connection handler (POC)
  ipcMain.handle('test-excel-connection', async () => {
    try {
      const result = await excel.testExcelConnection();
      return { success: true, message: result };
    } catch (error) {
      return {
        success: false,
        message: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  });

  // Write to Excel cell
  ipcMain.handle('excel:write', async (_, address, value) => {
    try {
      return excel.writeCell(address, value);
    } catch (error) {
      return {
        success: false,
        message: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  });

  // Format Excel cell
  ipcMain.handle('excel:format', async (_, address, options) => {
    try {
      return excel.formatCell(address, options);
    } catch (error) {
      return {
        success: false,
        message: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  });

  // Read active cell
  ipcMain.handle('excel:read', async () => {
    try {
      return excel.getActiveCell();
    } catch (error) {
      return {
        success: false,
        message: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  });

  // Inject VBA code
  ipcMain.handle('vba:inject', async (_, code) => {
    try {
      return excel.injectVBA(code);
    } catch (error) {
      return {
        success: false,
        message: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  });
}

module.exports = { registerHandlers };
