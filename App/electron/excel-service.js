const winax = require('winax');

let excelInstance = null;

function getExcel() {
  if (!excelInstance) {
    excelInstance = new winax.Object('Excel.Application', {
      activate: true,
      type: true
    });
    excelInstance.Visible = true;
  }
  return excelInstance;
}

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
    range.Font.Color = options.color; // BGR format
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

function injectVBA(code) {
  const excel = getExcel();
  const workbook = excel.ActiveWorkbook;
  const module = workbook.VBProject.VBComponents.Add(1); // 1 = vbext_ct_StdModule
  module.CodeModule.AddFromString(code);
  return { success: true };
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
    cellA1.Font.Color = 0x0000FF; // Red in BGR format

    return 'Successfully wrote "Hello World" to Excel cell A1!';
  } catch (error) {
    if (error instanceof Error) {
      throw new Error(`Excel COM Error: ${error.message}`);
    }
    throw error;
  }
}

module.exports = {
  getExcel,
  writeCell,
  formatCell,
  getActiveCell,
  injectVBA,
  testExcelConnection
};
