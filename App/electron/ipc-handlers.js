const { ipcMain, app } = require('electron');
const winax = require('winax');

// ============================================================================
// EXCEL COM CONNECTION
// ============================================================================

function getExcel() {
  // Use GetObject to connect to running Excel instance
  try {
    const excel = winax.GetObject('', 'Excel.Application');
    if (!excel) {
      throw new Error('Could not connect to Excel');
    }
    return excel;
  } catch (error) {
    throw new Error('Excel is not running. Please open Excel first.');
  }
}

// ============================================================================
// VBA INJECTION
// ============================================================================

function injectVBA(code) {
  try {
    // Step 1: Get running Excel instance
    let excel;
    try {
      excel = getExcel();
    } catch (error) {
      return {
        success: false,
        message: 'Excel is not running. Please open Excel with a workbook first.'
      };
    }

    // Step 2: Safety check - ensure workbook is open
    const workbook = excel.ActiveWorkbook;
    if (!workbook) {
      return {
        success: false,
        message: 'Please open a workbook first.'
      };
    }

    // Step 3: VBA Trust Check - wrap VBProject access
    let vbProject;
    try {
      vbProject = workbook.VBProject;
    } catch (error) {
      return {
        success: false,
        message: 'Action Blocked: Go to File > Options > Trust Center > Settings > Macro Settings > Check "Trust access to the VBA project object model".'
      };
    }

    const moduleName = 'MacroFlowModule';
    let existingModule = null;

    // Step 4: Check if MacroFlowModule exists
    try {
      const components = vbProject.VBComponents;
      for (let i = 1; i <= components.Count; i++) {
        const component = components.Item(i);
        if (component.Name === moduleName) {
          existingModule = component;
          break;
        }
      }
    } catch (error) {
      // Ignore iteration errors
    }

    // Step 5: Remove existing module if found
    if (existingModule) {
      try {
        vbProject.VBComponents.Remove(existingModule);
      } catch (error) {
        return {
          success: false,
          message: `Failed to remove existing module: ${error.message}`
        };
      }
    }

    // Step 6: Create new module
    const newModule = vbProject.VBComponents.Add(1); // 1 = vbext_ct_StdModule
    newModule.Name = moduleName;

    // Step 7: Add code to module
    const codeModule = newModule.CodeModule;
    const lineCount = codeModule.CountOfLines;
    codeModule.InsertLines(lineCount + 1, code);

    return {
      success: true,
      message: 'Module Created in Active Workbook'
    };
  } catch (error) {
    return {
      success: false,
      message: error instanceof Error ? error.message : 'Unknown error occurred'
    };
  }
}

// ============================================================================
// EXCEL CELL OPERATIONS (Legacy - for future features)
// ============================================================================

function writeCell(address, value) {
  const excel = getExcel();
  const range = excel.ActiveSheet.Range(address);
  range.Value = value;
  return { success: true, address, value };
}

function formatCell(address, options) {
  const excel = getExcel();
  const range = excel.ActiveSheet.Range(address);

  if (options.bold !== undefined) {
    range.Font.Bold = options.bold;
  }
  if (options.size !== undefined) {
    range.Font.Size = options.size;
  }
  if (options.color !== undefined) {
    range.Font.Color = options.color;
  }

  return { success: true, address };
}

function getActiveCell() {
  const excel = getExcel();
  const cell = excel.ActiveCell;
  return {
    address: cell.Address,
    value: cell.Value
  };
}

async function testExcelConnection() {
  try {
    const excel = getExcel();
    const workbook = excel.Workbooks.Add();
    const worksheet = workbook.ActiveSheet;
    const cellA1 = worksheet.Range('A1');

    cellA1.Value = 'Hello World';
    cellA1.Font.Bold = true;
    cellA1.Font.Size = 14;
    cellA1.Font.Color = 0x0000FF;

    return 'Successfully wrote "Hello World" to Excel cell A1!';
  } catch (error) {
    if (error instanceof Error) {
      throw new Error(`Excel COM Error: ${error.message}`);
    }
    throw error;
  }
}

// ============================================================================
// IPC HANDLERS REGISTRATION
// ============================================================================

function registerHandlers() {
  // Window control - Close app
  ipcMain.on('close-app', () => {
    app.quit();
  });

  // Test connection handler (POC)
  ipcMain.handle('test-excel-connection', async () => {
    try {
      const result = await testExcelConnection();
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
      return writeCell(address, value);
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
      return formatCell(address, options);
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
      return getActiveCell();
    } catch (error) {
      return {
        success: false,
        message: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  });

  // Inject VBA code - PRIMARY FEATURE
  ipcMain.handle('vba:inject', async (_, code) => {
    try {
      return injectVBA(code);
    } catch (error) {
      return {
        success: false,
        message: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  });
}

module.exports = { registerHandlers };
