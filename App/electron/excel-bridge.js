/**
 * ExcelBridge - Clean COM interface between Electron and Excel
 *
 * This is the ONLY file that should interact with Excel via COM.
 * All Excel operations go through this class.
 *
 * Usage:
 *   const excel = new ExcelBridge();
 *   const result = excel.injectModule('MacroFlowModule', vbaCode);
 *
 * Architecture:
 *   [React UI] → [IPC] → [ipc-handlers.js] → [ExcelBridge] → [Excel COM]
 */

const winax = require('winax');

// VBA Component Types (from vbext_ComponentType enum)
const VBA_COMPONENT = {
  STANDARD_MODULE: 1,  // vbext_ct_StdModule
  CLASS_MODULE: 2,     // vbext_ct_ClassModule
  FORM: 3,             // vbext_ct_MSForm
  DOCUMENT: 100        // vbext_ct_Document (ThisWorkbook, Sheet1, etc.)
};

class ExcelBridge {

  // ============================================================================
  // CONNECTION
  // ============================================================================

  /**
   * Connect to the running Excel instance
   * Uses GetObject to attach to existing Excel - never creates new instances
   * @returns {object} Excel.Application COM object
   * @throws {Error} If Excel is not running
   */
  getApp() {
    try {
      const excel = winax.GetObject('', 'Excel.Application');
      if (!excel) {
        throw new Error('Excel application not found');
      }
      return excel;
    } catch (error) {
      throw new Error('NO_EXCEL: Excel is not running. Please open Excel first.');
    }
  }

  /**
   * Get the active workbook
   * @returns {object} Workbook COM object
   * @throws {Error} If no workbook is open
   */
  getActiveWorkbook() {
    const excel = this.getApp();
    const workbook = excel.ActiveWorkbook;
    if (!workbook) {
      throw new Error('NO_WORKBOOK: No workbook is open. Please open or create a workbook.');
    }
    return workbook;
  }

  /**
   * Get the VBA project from the active workbook
   * @returns {object} VBProject COM object
   * @throws {Error} If VBA access is not trusted
   */
  getVBProject() {
    const workbook = this.getActiveWorkbook();
    try {
      const vbProject = workbook.VBProject;
      if (!vbProject) {
        throw new Error('VBA project not accessible');
      }
      return vbProject;
    } catch (error) {
      throw new Error(
        'VBA_BLOCKED: Trust access to VBA project is disabled. ' +
        'Go to File → Options → Trust Center → Trust Center Settings → ' +
        'Macro Settings → Check "Trust access to the VBA project object model".'
      );
    }
  }

  // ============================================================================
  // VBA MODULE OPERATIONS
  // ============================================================================

  /**
   * Inject or replace a VBA module in the active workbook
   * @param {string} moduleName - Name of the module (e.g., 'MacroFlowModule')
   * @param {string} code - VBA code to inject
   * @returns {{ success: boolean, message: string }}
   */
  injectModule(moduleName, code) {
    try {
      const vbProject = this.getVBProject();

      // Remove existing module if it exists
      this._removeModule(vbProject, moduleName);

      // Create new module
      const newModule = vbProject.VBComponents.Add(VBA_COMPONENT.STANDARD_MODULE);
      newModule.Name = moduleName;

      // Insert code
      const codeModule = newModule.CodeModule;
      codeModule.InsertLines(codeModule.CountOfLines + 1, code);

      return { success: true, message: `Module "${moduleName}" created successfully` };
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * Remove a module from the VBA project
   * @param {object} vbProject - VBProject COM object
   * @param {string} moduleName - Name of module to remove
   * @private
   */
  _removeModule(vbProject, moduleName) {
    try {
      const components = vbProject.VBComponents;
      for (let i = 1; i <= components.Count; i++) {
        const component = components.Item(i);
        if (component.Name === moduleName) {
          vbProject.VBComponents.Remove(component);
          return;
        }
      }
    } catch (error) {
      // Module doesn't exist or can't be removed - that's fine
    }
  }

  /**
   * Run a VBA macro by name
   * @param {string} macroName - Name of macro to run (e.g., 'MacroFlowModule.MyMacro')
   * @returns {{ success: boolean, message: string }}
   */
  runMacro(macroName) {
    try {
      const excel = this.getApp();
      excel.Run(macroName);
      return { success: true, message: `Executed "${macroName}"` };
    } catch (error) {
      return { success: false, message: `Failed to run macro: ${error.message}` };
    }
  }

  // ============================================================================
  // CELL OPERATIONS
  // ============================================================================

  /**
   * Read a cell value
   * @param {string} address - Cell address (e.g., 'A1', 'B2:C5')
   * @returns {{ success: boolean, value: any, address: string }}
   */
  readCell(address) {
    try {
      const excel = this.getApp();
      const range = excel.ActiveSheet.Range(address);
      return { success: true, address, value: range.Value };
    } catch (error) {
      return { success: false, address, message: error.message };
    }
  }

  /**
   * Write a value to a cell
   * @param {string} address - Cell address
   * @param {any} value - Value to write
   * @returns {{ success: boolean, address: string }}
   */
  writeCell(address, value) {
    try {
      const excel = this.getApp();
      const range = excel.ActiveSheet.Range(address);
      range.Value = value;
      return { success: true, address };
    } catch (error) {
      return { success: false, address, message: error.message };
    }
  }

  /**
   * Get the currently selected cell
   * @returns {{ success: boolean, address: string, value: any }}
   */
  getSelection() {
    try {
      const excel = this.getApp();
      const cell = excel.ActiveCell;
      return {
        success: true,
        address: cell.Address.replace(/\$/g, ''), // Remove $ signs
        value: cell.Value
      };
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  // ============================================================================
  // WORKBOOK INFO
  // ============================================================================

  /**
   * Get info about the active workbook
   * @returns {{ success: boolean, name: string, path: string, sheets: string[] }}
   */
  getWorkbookInfo() {
    try {
      const workbook = this.getActiveWorkbook();
      const sheets = [];
      for (let i = 1; i <= workbook.Sheets.Count; i++) {
        sheets.push(workbook.Sheets.Item(i).Name);
      }
      return {
        success: true,
        name: workbook.Name,
        path: workbook.FullName,
        sheets
      };
    } catch (error) {
      return { success: false, message: error.message };
    }
  }
}

// Export singleton instance
module.exports = new ExcelBridge();
