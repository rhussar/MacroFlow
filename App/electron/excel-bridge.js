/**
 * ExcelBridge - COM bridge for Excel automation.
 *
 * This is the only file that should touch Excel via COM.
 * All UI/IPC layers call into this class.
 */

const winax = require('winax');

// VBA Component Types (vbext_ComponentType)
const VBA_COMPONENT_TYPE = {
  STANDARD_MODULE: 1,
  CLASS_MODULE: 2,
  FORM: 3,
  DOCUMENT: 100
};

const VBA_COMPONENT_NAME = {
  [VBA_COMPONENT_TYPE.STANDARD_MODULE]: 'Standard Module',
  [VBA_COMPONENT_TYPE.CLASS_MODULE]: 'Class Module',
  [VBA_COMPONENT_TYPE.FORM]: 'UserForm',
  [VBA_COMPONENT_TYPE.DOCUMENT]: 'Document'
};

class ExcelBridge {
  constructor() {
    this._proceduresCache = null;
    this._cachedApp = null;
    this._cachedAppActivate = null;
  }

  // ===========================================================================
  // CONNECTION
  // ===========================================================================

  /**
   * Connect to a running Excel instance (never creates a new one).
   * Caches the COM reference and reuses it until the connection goes stale.
   * @returns {object} Excel.Application COM object
   * @throws {Error} If Excel is not running
   */
  getApp(options = {}) {
    const { activate = true } = options;

    // Try to reuse the cached COM reference if the activate mode matches
    if (this._cachedApp && this._cachedAppActivate === activate) {
      try {
        // Probe the cached reference — if Excel was closed this will throw
        void this._cachedApp.Version;
        return this._cachedApp;
      } catch {
        // Stale reference — Excel was closed or restarted
        this._cachedApp = null;
        this._cachedAppActivate = null;
      }
    }

    try {
      const excel = new winax.Object('Excel.Application', { activate });
      if (!excel) {
        throw new Error('Excel application not found');
      }
      this._cachedApp = excel;
      this._cachedAppActivate = activate;
      return excel;
    } catch (error) {
      this._cachedApp = null;
      this._cachedAppActivate = null;
      throw new Error('NO_EXCEL: Excel is not running. Please open Excel first.');
    }
  }

  /**
   * Get the active workbook.
   * @returns {object} Workbook COM object
   * @throws {Error} If no workbook is open
   */
  getActiveWorkbook(options = {}) {
    const excel = this.getApp(options);
    const workbook = excel.ActiveWorkbook;
    if (!workbook) {
      throw new Error('NO_WORKBOOK: No workbook is open. Please open or create a workbook.');
    }
    return workbook;
  }

  /**
   * Get the VBA project from the active workbook.
   * @returns {object} VBProject COM object
   * @throws {Error} If VBA access is not trusted
   */
  getVBProject(options = {}) {
    const workbook = this.getActiveWorkbook(options);
    return this._getVBProjectForWorkbook(workbook);
  }

  _getVBProjectForWorkbook(workbook) {
    try {
      const vbProject = workbook.VBProject;
      if (!vbProject) {
        throw new Error('VBA project not accessible');
      }
      return vbProject;
    } catch (error) {
      throw new Error(
        'VBA_BLOCKED: Trust access to VBA project is disabled. ' +
          'Go to File > Options > Trust Center > Trust Center Settings > ' +
          'Macro Settings > Check "Trust access to the VBA project object model".'
      );
    }
  }

  // ===========================================================================
  // HELPERS (PURE)
  // ===========================================================================

  _getWorkbookCacheKey(workbook) {
    const path = workbook && workbook.FullName ? String(workbook.FullName) : '';
    const name = workbook && workbook.Name ? String(workbook.Name) : '';
    return `${path}::${name}`;
  }

  _describeWorkbook(workbook) {
    return {
      name: workbook.Name,
      path: workbook.FullName
    };
  }

  _findOpenWorkbookByName(excel, workbookName) {
    const normalizedTarget = String(workbookName || '').trim().toLowerCase();
    if (!normalizedTarget) {
      return null;
    }

    const workbooks = excel?.Workbooks;
    const count = workbooks ? Number(workbooks.Count) : 0;
    if (!Number.isFinite(count) || count < 1) {
      return null;
    }

    for (let i = 1; i <= count; i++) {
      const workbook = workbooks.Item(i);
      if (!workbook) {
        continue;
      }

      const candidateName = String(workbook.Name || '').trim().toLowerCase();
      if (candidateName === normalizedTarget) {
        return workbook;
      }
    }

    return null;
  }

  _findComponentByName(vbProject, name) {
    const components = vbProject.VBComponents;
    for (let i = 1; i <= components.Count; i++) {
      const component = components.Item(i);
      if (component.Name === name) {
        return component;
      }
    }
    return null;
  }

  _qualifyWorkbookName(name) {
    const safe = String(name || '');
    if (!safe) {
      return '';
    }
    if (/\s/.test(safe)) {
      return `'${safe.replace(/'/g, "''")}'`;
    }
    return safe;
  }

  _normalizeMacroNameForWorkbook(macroName, workbook, procedures = []) {
    if (!macroName) {
      return '';
    }

    const trimmed = String(macroName).trim();
    if (!trimmed) {
      return '';
    }

    if (trimmed.includes('!')) {
      return trimmed;
    }

    const workbookName = this._qualifyWorkbookName(workbook?.Name);
    if (!workbookName) {
      return trimmed;
    }

    if (trimmed.includes('.')) {
      return `${workbookName}!${trimmed}`;
    }

    const sourceProcedures = Array.isArray(procedures) ? procedures : [];
    const matches = sourceProcedures.filter(
      (proc) => String(proc?.name || '').toLowerCase() === trimmed.toLowerCase()
    );
    if (matches.length === 1) {
      return `${workbookName}!${matches[0].module}.${matches[0].name}`;
    }

    return trimmed;
  }

  _normalizeMacroName(macroName, workbook) {
    const { procedures } = this._listMacroProcedures();
    return this._normalizeMacroNameForWorkbook(macroName, workbook, procedures);
  }

  _componentTypeName(typeId) {
    return VBA_COMPONENT_NAME[typeId] || `Unknown (${typeId})`;
  }

  _listModulesForWorkbook(workbook, vbProject) {
    const components = vbProject.VBComponents;
    const modules = [];

    for (let i = 1; i <= components.Count; i++) {
      const component = components.Item(i);
      const codeModule = component.CodeModule;
      const lineCount = codeModule ? codeModule.CountOfLines : 0;
      modules.push({
        name: component.Name,
        typeId: component.Type,
        type: this._componentTypeName(component.Type),
        lineCount
      });
    }

    return modules;
  }

  _listProceduresForWorkbook(workbook, vbProject) {
    const components = vbProject.VBComponents;
    const workbookKey = this._getWorkbookCacheKey(workbook);
    const previousModuleEntries =
      this._proceduresCache && this._proceduresCache.workbookKey === workbookKey
        ? this._proceduresCache.moduleEntries
        : {};
    const nextModuleEntries = {};
    const procedures = [];

    for (let i = 1; i <= components.Count; i++) {
      const component = components.Item(i);
      const moduleName = String(component.Name);
      const codeModule = component.CodeModule;
      if (!codeModule) {
        continue;
      }

      const lineCount = Number(codeModule.CountOfLines) || 0;
      if (lineCount < 1) {
        continue;
      }

      const cachedModule = previousModuleEntries[moduleName];
      if (cachedModule && cachedModule.lineCount === lineCount) {
        nextModuleEntries[moduleName] = cachedModule;
        cachedModule.procedures.forEach((proc) => {
          procedures.push({ ...proc });
        });
        continue;
      }

      const codeText = codeModule.Lines(1, lineCount);
      const parsed = this._parseProcedures(codeText);
      const moduleProcedures = parsed.map((proc) => ({
        module: moduleName,
        ...proc
      }));

      nextModuleEntries[moduleName] = {
        lineCount,
        procedures: moduleProcedures
      };

      moduleProcedures.forEach((proc) => {
        procedures.push({ ...proc });
      });
    }

    this._proceduresCache = {
      workbookKey,
      moduleEntries: nextModuleEntries
    };

    return procedures;
  }

  _parseProcedures(codeText) {
    const procedures = [];
    const regex =
      /^\s*(Public|Private|Friend)?\s*(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)/gim;
    let match;
    while ((match = regex.exec(codeText)) !== null) {
      const scope = match[1] ? match[1] : 'Implicit';
      const isStatic = Boolean(match[2]);
      const kind = match[3].replace(/\s+/g, ' ');
      const name = match[5];
      procedures.push({
        name,
        kind,
        scope,
        isStatic
      });
    }
    return procedures;
  }

  _normalizeValue(value) {
    if (value === null || value === undefined) {
      return null;
    }
    if (value instanceof Date) {
      return value.toISOString();
    }
    if (Array.isArray(value)) {
      return value.map((item) => this._normalizeValue(item));
    }
    if (typeof value === 'object') {
      if (typeof value.valueOf === 'function') {
        const primitive = value.valueOf();
        if (primitive === null || primitive === undefined) {
          return null;
        }
        if (typeof primitive !== 'object') {
          return primitive;
        }
      }
      return String(value);
    }
    return value;
  }

  _columnLetter(index) {
    let dividend = index;
    let columnName = '';
    while (dividend > 0) {
      const modulo = (dividend - 1) % 26;
      columnName = String.fromCharCode(65 + modulo) + columnName;
      dividend = Math.floor((dividend - modulo) / 26);
    }
    return columnName;
  }

  _classifyValue(value) {
    if (value === null || value === undefined || value === '') {
      return 'empty';
    }
    if (typeof value === 'number' && Number.isFinite(value)) {
      return 'number';
    }
    if (typeof value === 'boolean') {
      return 'boolean';
    }
    if (typeof value === 'string') {
      const trimmed = value.trim();
      if (!trimmed) {
        return 'empty';
      }
      if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
        return 'number-text';
      }
      if (!Number.isNaN(Date.parse(trimmed))) {
        return 'date-text';
      }
      return 'text';
    }
    return 'other';
  }

  _normalizeRangeValues(values) {
    const normalizeRow = (row) => {
      if (!Array.isArray(row)) {
        return [this._normalizeValue(row)];
      }
      const trimmedRow = row.length > 0 && row[0] === undefined ? row.slice(1) : row;
      return trimmedRow.map((cell) => this._normalizeValue(cell));
    };

    if (Array.isArray(values)) {
      const hasRowArrays = Array.isArray(values[0]) || Array.isArray(values[1]);
      if (hasRowArrays) {
        const rows = values.length > 0 && values[0] === undefined ? values.slice(1) : values;
        return rows.map((row) => normalizeRow(row));
      }
      return [normalizeRow(values)];
    }

    return [[this._normalizeValue(values)]];
  }

  _getRuntimeModuleCode() {
    return [
      'Option Explicit',
      '',
      'Public Function MacroFlow_RunMacro(ByVal macroName As String) As String',
      '    On Error GoTo Handler',
      '    Application.Run macroName',
      '    MacroFlow_RunMacro = "OK"',
      '    Exit Function',
      'Handler:',
      '    MacroFlow_RunMacro = "ERR|" & Err.Number & "|" & Err.Description & "|" & Err.Source',
      'End Function'
    ].join('\n');
  }

  _ensureRuntimeModule() {
    const vbProject = this.getVBProject();
    const moduleName = 'MacroFlow_Runtime';
    const code = this._getRuntimeModuleCode();
    const existing = this._findComponentByName(vbProject, moduleName);

    if (existing) {
      const codeModule = existing.CodeModule;
      if (codeModule.CountOfLines > 0) {
        const existingCode = codeModule.Lines(1, codeModule.CountOfLines);
        if (existingCode.includes('MacroFlow_RunMacro')) {
          return;
        }
        codeModule.DeleteLines(1, codeModule.CountOfLines);
      }
      codeModule.AddFromString(code);
      return;
    }

    const newModule = vbProject.VBComponents.Add(VBA_COMPONENT_TYPE.STANDARD_MODULE);
    newModule.Name = moduleName;
    newModule.CodeModule.AddFromString(code);
  }

  _withWorkbookAtPath(path, callback) {
    const excel = this.getApp();
    const state = {
      screenUpdating: excel.ScreenUpdating,
      displayAlerts: excel.DisplayAlerts,
      enableEvents: excel.EnableEvents
    };
    let workbook = null;
    let previousWorkbook = null;

    try {
      previousWorkbook = excel.ActiveWorkbook;
      excel.ScreenUpdating = false;
      excel.DisplayAlerts = false;
      excel.EnableEvents = false;

      workbook = excel.Workbooks.Open(path, 0, true);
      try {
        workbook.Windows.Item(1).Visible = false;
      } catch (error) {
        // Ignore if window visibility cannot be changed.
      }

      return callback(workbook, excel);
    } finally {
      if (workbook) {
        try {
          workbook.Close(false);
        } catch (error) {
          // Ignore close errors.
        }
      }
      try {
        excel.ScreenUpdating = state.screenUpdating;
        excel.DisplayAlerts = state.displayAlerts;
        excel.EnableEvents = state.enableEvents;
      } catch (error) {
        // Ignore restore errors.
      }
      if (previousWorkbook) {
        try {
          previousWorkbook.Activate();
        } catch (error) {
          // Ignore activation errors.
        }
      }
    }
  }

  _prepareMacroRun(excel, workbook) {
    try {
      workbook.Activate();
    } catch (error) {
      // Ignore if activation fails.
    }

    try {
      if (excel.ActiveWindow) {
        excel.ActiveWindow.Activate();
      }
    } catch (error) {
      // Ignore window activation issues.
    }

    let ready = true;
    try {
      ready = Boolean(excel.Ready);
    } catch (error) {
      ready = true;
    }

    return {
      ready,
      message: ready ? '' : 'Excel is busy or in edit mode. Finish editing a cell to run macros.'
    };
  }

  _buildSelectionContext(excel) {
    let selectionAddress = '';
    let activeCellAddress = '';
    let activeCellValue = null;
    let selectionInTable = false;
    let tableInfo = null;

    try {
      const selection = excel.Selection;
      if (selection) {
        selectionAddress = String(selection.Address).replace(/\$/g, '');
        try {
          const listObject = selection.ListObject;
          if (listObject) {
            selectionInTable = true;
            tableInfo = {
              name: String(listObject.Name),
              range: String(listObject.Range.Address).replace(/\$/g, '')
            };
          }
        } catch (error) {
          // Not in a table.
        }
      }
    } catch (error) {
      // Ignore selection errors.
    }

    try {
      const activeCell = excel.ActiveCell;
      if (activeCell) {
        activeCellAddress = String(activeCell.Address).replace(/\$/g, '');
        activeCellValue = this._normalizeValue(activeCell.Value2);
        if (!selectionInTable) {
          try {
            const listObject = activeCell.ListObject;
            if (listObject) {
              selectionInTable = true;
              tableInfo = {
                name: String(listObject.Name),
                range: String(listObject.Range.Address).replace(/\$/g, '')
              };
            }
          } catch (error) {
            // Not in a table.
          }
        }
      }
    } catch (error) {
      // Ignore active cell errors.
    }

    return {
      address: selectionAddress,
      activeCell: {
        address: activeCellAddress,
        value: activeCellValue
      },
      inTable: selectionInTable,
      table: tableInfo
    };
  }

  _collectWorksheetMetadata({ excel, workbook, sheet, includeSelection }) {
    const usedRange = sheet.UsedRange;
    const startRow = usedRange.Row;
    const startColumn = usedRange.Column;
    const totalRows = Number(usedRange.Rows.Count);
    const totalColumns = Number(usedRange.Columns.Count);

    const SAMPLE_ROWS = 50;
    const SAMPLE_COLS = 25;

    const previewRows = Math.min(totalRows, SAMPLE_ROWS);
    const previewCols = Math.min(totalColumns, SAMPLE_COLS);

    const previewRange = sheet.Range(
      sheet.Cells(startRow, startColumn),
      sheet.Cells(startRow + previewRows - 1, startColumn + previewCols - 1)
    );
    const previewValues = previewRange.Value2;
    const preview = this._normalizeRangeValues(previewValues);

    const headerRange = sheet.Range(
      sheet.Cells(startRow, startColumn),
      sheet.Cells(startRow, startColumn + previewCols - 1)
    );
    const headerValues = this._normalizeRangeValues(headerRange.Value2);
    const headers = headerValues[0] || [];

    const headerAddressMap = {};
    for (let colIndex = 0; colIndex < previewCols; colIndex++) {
      const columnNumber = startColumn + colIndex;
      const address = `${this._columnLetter(columnNumber)}${startRow}`;
      headerAddressMap[address] = headers[colIndex] ?? null;
    }

    let formulaGrid = [];
    try {
      formulaGrid = this._normalizeRangeValues(previewRange.Formula);
    } catch (error) {
      formulaGrid = [];
    }

    const sampleRowCount = Math.max(0, Math.min(4, totalRows - 1));
    let sampleRows = [];
    let sampleRowStart = startRow + 1;
    if (sampleRowCount > 0) {
      const sampleRange = sheet.Range(
        sheet.Cells(sampleRowStart, startColumn),
        sheet.Cells(sampleRowStart + sampleRowCount - 1, startColumn + previewCols - 1)
      );
      sampleRows = this._normalizeRangeValues(sampleRange.Value2);
    }

    const columns = [];
    for (let colIndex = 0; colIndex < previewCols; colIndex++) {
      const columnNumber = startColumn + colIndex;
      const columnLetter = this._columnLetter(columnNumber);
      const header = headers[colIndex] !== undefined ? headers[colIndex] : '';

      const typeCounts = {
        empty: 0,
        number: 0,
        boolean: 0,
        text: 0,
        'number-text': 0,
        'date-text': 0,
        other: 0
      };
      const examples = [];
      const exampleSet = new Set();
      let nonEmpty = 0;
      let numericMin = null;
      let numericMax = null;
      let formulaCells = 0;

      const analysisRows = sampleRows.length ? sampleRows : preview.slice(1);
      for (let rowIndex = 0; rowIndex < analysisRows.length; rowIndex++) {
        const value = analysisRows[rowIndex]?.[colIndex];
        const type = this._classifyValue(value);
        typeCounts[type] += 1;
        if (type !== 'empty') {
          nonEmpty += 1;
          const example = value === null || value === undefined ? '' : String(value);
          if (example && !exampleSet.has(example) && examples.length < 6) {
            examples.push(example);
            exampleSet.add(example);
          }
        }
        if (type === 'number') {
          if (numericMin === null || value < numericMin) numericMin = value;
          if (numericMax === null || value > numericMax) numericMax = value;
        }

        const formulaCell = formulaGrid[rowIndex + 1]?.[colIndex];
        if (typeof formulaCell === 'string' && formulaCell.startsWith('=')) {
          formulaCells += 1;
        }
      }

      const typeSummaryParts = Object.entries(typeCounts)
        .filter(([key, count]) => key !== 'empty' && count > 0)
        .sort((a, b) => b[1] - a[1])
        .map(([key, count]) => `${key} ${count}`);
      if (formulaCells > 0) {
        typeSummaryParts.push(`formulas ${formulaCells}`);
      }
      const typeSummary = typeSummaryParts.length ? typeSummaryParts.join(' | ') : 'empty';

      columns.push({
        index: columnNumber,
        column: columnLetter,
        header,
        nonEmpty,
        empty: typeCounts.empty,
        typeSummary,
        examples,
        numericMin,
        numericMax,
        formulaCells
      });
    }

    const sheetNames = [];
    let activeSheetName = '';
    for (let i = 1; i <= workbook.Sheets.Count; i++) {
      const sheetItem = workbook.Sheets.Item(i);
      sheetNames.push(String(sheetItem.Name));
    }
    if (workbook.ActiveSheet) {
      activeSheetName = String(workbook.ActiveSheet.Name);
    }

    const selectionContext = includeSelection ? this._buildSelectionContext(excel) : null;

    const structuralContext = {
      workbook: this._describeWorkbook(workbook),
      sheets: sheetNames,
      activeSheet: activeSheetName
    };

    const dataContext = {
      usedRange: {
        address: String(usedRange.Address),
        startRow,
        startColumn,
        rows: totalRows,
        columns: totalColumns
      },
      headers,
      headersByAddress: headerAddressMap,
      sampleRows: {
        startRow: sampleRowCount ? sampleRowStart : null,
        endRow: sampleRowCount ? sampleRowStart + sampleRowCount - 1 : null,
        rows: sampleRows
      },
      columns
    };

    const llmContext = {
      structural: {
        workbookName: structuralContext.workbook.name,
        workbookPath: structuralContext.workbook.path,
        worksheetNames: structuralContext.sheets,
        activeSheet: structuralContext.activeSheet
      },
      data: {
        usedRange: dataContext.usedRange,
        headersByAddress: dataContext.headersByAddress,
        sampleRows: dataContext.sampleRows,
        columns: dataContext.columns.map((column) => ({
          column: column.column,
          header: column.header,
          typeSummary: column.typeSummary,
          nonEmpty: column.nonEmpty,
          numericMin: column.numericMin,
          numericMax: column.numericMax,
          examples: column.examples,
          formulaCells: column.formulaCells
        }))
      },
      selection: selectionContext
    };

    return {
      sheet: {
        name: String(sheet.Name),
        index: Number(sheet.Index)
      },
      structuralContext,
      dataContext,
      selectionContext,
      llmContext,
      preview: {
        rows: preview,
        truncated: totalRows > previewRows || totalColumns > previewCols
      }
    };
  }

  // ===========================================================================
  // WORKBOOK INFO
  // ===========================================================================

  /**
   * Get info about the active workbook.
   * @returns {{ success: boolean, name: string, path: string, activeSheet: string, sheets: string[] }}
   */
  getWorkbookInfo() {
    try {
      const workbook = this.getActiveWorkbook();
      let activeSheet = '';
      try {
        activeSheet = workbook.ActiveSheet ? String(workbook.ActiveSheet.Name) : '';
      } catch (error) {
        activeSheet = '';
      }
      const sheets = [];
      for (let i = 1; i <= workbook.Sheets.Count; i++) {
        sheets.push(workbook.Sheets.Item(i).Name);
      }
      return {
        success: true,
        name: workbook.Name,
        path: workbook.FullName,
        activeSheet,
        sheets
      };
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * Get all open workbooks in Excel.
   * @returns {{ success: boolean, workbooks: Array<{ name: string, path: string }>, message?: string }}
   */
  getOpenWorkbooks() {
    try {
      const excel = this.getApp();
      const workbooks = [];
      const count = excel.Workbooks.Count;

      for (let i = 1; i <= count; i++) {
        const workbook = excel.Workbooks.Item(i);
        workbooks.push(this._describeWorkbook(workbook));
      }

      return { success: true, workbooks };
    } catch (error) {
      if (error.message.includes('NO_EXCEL')) {
        return { success: false, workbooks: [], message: 'Excel not found' };
      }
      return { success: false, workbooks: [], message: error.message };
    }
  }

  /**
   * List worksheets in the active workbook with basic UsedRange stats.
   * @returns {{ success: boolean, workbook?: { name: string, path: string }, sheets: Array }}
   */
  listWorksheets() {
    try {
      const workbook = this.getActiveWorkbook();
      const activeSheetName = workbook.ActiveSheet ? String(workbook.ActiveSheet.Name) : '';
      const sheets = [];
      for (let i = 1; i <= workbook.Sheets.Count; i++) {
        const sheet = workbook.Sheets.Item(i);
        let usedRange = null;
        try {
          usedRange = sheet.UsedRange;
        } catch (error) {
          usedRange = null;
        }
        sheets.push({
          name: String(sheet.Name),
          index: i,
          visible: Number(sheet.Visible),
          active: activeSheetName ? activeSheetName === String(sheet.Name) : false,
          usedRange: usedRange
            ? {
                address: String(usedRange.Address),
                rows: Number(usedRange.Rows.Count),
                columns: Number(usedRange.Columns.Count)
              }
            : null
        });
      }

      return {
        success: true,
        workbook: this._describeWorkbook(workbook),
        sheets
      };
    } catch (error) {
      return { success: false, sheets: [], message: error.message };
    }
  }

  /**
   * Build workbook + sheet metadata for LLM prompting.
   * @param {{ sheetName?: string }} options
   * @returns {{ success: boolean, sheet?: object, structuralContext?: object, dataContext?: object, selectionContext?: object, llmContext?: object }}
   */
  getWorksheetMetadata(options = {}) {
    try {
      const excel = this.getApp();
      const workbook = this.getActiveWorkbook();
      const sheet = options.sheetName
        ? workbook.Sheets.Item(options.sheetName)
        : workbook.ActiveSheet;

      const metadata = this._collectWorksheetMetadata({
        excel,
        workbook,
        sheet,
        includeSelection: true
      });

      return {
        success: true,
        message: `Metadata read for "${sheet.Name}"`,
        ...metadata
      };
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * Build metadata for a closed workbook path (opened read-only in the background).
   * @param {{ path: string, sheetName?: string }} options
   * @returns {{ success: boolean, sheet?: object, structuralContext?: object, dataContext?: object, llmContext?: object }}
   */
  getClosedWorkbookMetadata(options = {}) {
    if (!options.path) {
      return { success: false, message: 'Missing workbook path.' };
    }

    try {
      return this._withWorkbookAtPath(options.path, (workbook, excel) => {
        const sheet = options.sheetName
          ? workbook.Sheets.Item(options.sheetName)
          : workbook.ActiveSheet;
        const metadata = this._collectWorksheetMetadata({
          excel,
          workbook,
          sheet,
          includeSelection: false
        });
        return {
          success: true,
          message: `Metadata read for "${sheet.Name}"`,
          ...metadata
        };
      });
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  // ===========================================================================
  // CELL OPERATIONS (LEGACY)
  // ===========================================================================

  /**
   * Read a cell value.
   * @param {string} address - Cell address (e.g., 'A1', 'B2:C5')
   * @returns {{ success: boolean, value: any, address: string }}
   */
  readCell(address) {
    try {
      const excel = this.getApp();
      const range = excel.ActiveSheet.Range(address);
      return { success: true, address, value: this._normalizeValue(range.Value2) };
    } catch (error) {
      return { success: false, address, message: error.message };
    }
  }

  /**
   * Write a value to a cell.
   * @param {string} address - Cell address
   * @param {any} value - Value to write
   * @returns {{ success: boolean, address: string }}
   */
  writeCell(address, value) {
    try {
      const excel = this.getApp();
      const range = excel.ActiveSheet.Range(address);
      range.Value2 = value;
      return { success: true, address };
    } catch (error) {
      return { success: false, address, message: error.message };
    }
  }

  /**
   * Get the currently selected cell.
   * @returns {{ success: boolean, address: string, value: any }}
   */
  getSelection() {
    try {
      const excel = this.getApp();
      const cell = excel.ActiveCell;
      return {
        success: true,
        address: String(cell.Address).replace(/\$/g, ''),
        value: this._normalizeValue(cell.Value2)
      };
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * Highlight the current selection with a color.
   * @param {string} colorName - 'Yellow', 'Green', 'Red', or 'None'
   * @returns {{ success: boolean, color: string }}
   */
  highlightSelection(colorName) {
    const colorMap = {
      Yellow: 65535,
      Green: 5296274,
      Red: 255,
      None: -4142
    };

    try {
      const excel = this.getApp();
      const selection = excel.Selection;

      if (!selection) {
        return { success: false, message: 'No cells selected' };
      }

      const colorValue = colorMap[colorName];
      if (colorValue === undefined) {
        return { success: false, message: `Unknown color: ${colorName}` };
      }

      if (colorName === 'None') {
        selection.Interior.ColorIndex = -4142;
      } else {
        selection.Interior.Color = colorValue;
      }

      return { success: true, color: colorName };
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  // ===========================================================================
  // VBA MODULES & PROCEDURES
  // ===========================================================================

  /**
   * List modules in the active workbook's VBA project.
   * @returns {{ success: boolean, workbook?: { name: string, path: string }, modules: Array }}
   */
  listModules(options = {}) {
    const { activate = true } = options;
    try {
      const workbook = this.getActiveWorkbook({ activate });
      const vbProject = this._getVBProjectForWorkbook(workbook);
      const modules = this._listModulesForWorkbook(workbook, vbProject);

      return {
        success: true,
        workbook: this._describeWorkbook(workbook),
        modules
      };
    } catch (error) {
      return { success: false, modules: [], message: error.message };
    }
  }

  /**
   * List modules in a specific open workbook's VBA project.
   * @param {string} workbookName
   * @param {{ activate?: boolean }} options
   * @returns {{ success: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, modules: Array, message?: string }}
   */
  listModulesByWorkbookName(workbookName, options = {}) {
    const { activate = true } = options;
    const normalizedName = String(workbookName || '').trim();
    if (!normalizedName) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        modules: [],
        message: 'Workbook name is required.'
      };
    }

    try {
      const excel = this.getApp({ activate });
      const workbook = this._findOpenWorkbookByName(excel, normalizedName);
      if (!workbook) {
        return {
          success: true,
          workbookFound: false,
          workbook: null,
          modules: [],
          message: `Workbook "${normalizedName}" is not open.`
        };
      }

      const vbProject = this._getVBProjectForWorkbook(workbook);
      const modules = this._listModulesForWorkbook(workbook, vbProject);

      return {
        success: true,
        workbookFound: true,
        workbook: this._describeWorkbook(workbook),
        modules
      };
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        modules: [],
        message: error.message
      };
    }
  }

  /**
   * List procedures (Subs/Functions/Properties) in the active workbook.
   * @returns {{ success: boolean, workbook?: { name: string, path: string }, procedures: Array }}
   */
  listProcedures(options = {}) {
    const { activate = true } = options;
    try {
      const workbook = this.getActiveWorkbook({ activate });
      const vbProject = this._getVBProjectForWorkbook(workbook);
      const procedures = this._listProceduresForWorkbook(workbook, vbProject);

      return {
        success: true,
        workbook: this._describeWorkbook(workbook),
        procedures
      };
    } catch (error) {
      this._proceduresCache = null;
      return { success: false, procedures: [], message: error.message };
    }
  }

  /**
   * List procedures (Subs/Functions/Properties) for a specific open workbook name.
   * @param {string} workbookName
   * @param {{ activate?: boolean }} options
   * @returns {{ success: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, procedures: Array, message?: string }}
   */
  listProceduresByWorkbookName(workbookName, options = {}) {
    const { activate = true } = options;
    const normalizedName = String(workbookName || '').trim();
    if (!normalizedName) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        procedures: [],
        message: 'Workbook name is required.'
      };
    }

    try {
      const excel = this.getApp({ activate });
      const workbook = this._findOpenWorkbookByName(excel, normalizedName);
      if (!workbook) {
        return {
          success: true,
          workbookFound: false,
          workbook: null,
          procedures: [],
          message: `Workbook "${normalizedName}" is not open.`
        };
      }

      const vbProject = this._getVBProjectForWorkbook(workbook);
      const procedures = this._listProceduresForWorkbook(workbook, vbProject);

      return {
        success: true,
        workbookFound: true,
        workbook: this._describeWorkbook(workbook),
        procedures
      };
    } catch (error) {
      this._proceduresCache = null;
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        procedures: [],
        message: error.message
      };
    }
  }

  /**
   * Inject or replace a VBA module in the active workbook.
   * @param {string} moduleName - Name of the module (e.g., 'MacroFlowModule')
   * @param {string} code - VBA code to inject
   * @returns {{ success: boolean, message: string }}
   */
  injectModule(moduleName, code) {
    try {
      const vbProject = this.getVBProject();

      this._removeModule(vbProject, moduleName);

      const newModule = vbProject.VBComponents.Add(VBA_COMPONENT_TYPE.STANDARD_MODULE);
      newModule.Name = moduleName;

      const codeModule = newModule.CodeModule;
      codeModule.InsertLines(codeModule.CountOfLines + 1, code);

      return { success: true, message: `Module "${moduleName}" created successfully` };
    } catch (error) {
      return { success: false, message: error.message };
    }
  }

  /**
   * Remove a module from the VBA project.
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
      // Module doesn't exist or can't be removed.
    }
  }

  /**
   * Run a VBA macro by name.
   * @param {string} macroName - e.g., 'ModuleName.MacroName' or 'MacroName'
   * @returns {{ success: boolean, message: string }}
   */
  runMacro(macroName) {
    return this.runMacroWithTrap(macroName);
  }

  // ===========================================================================
  // RUNTIME + SHORTCUTS
  // ===========================================================================

  /**
   * Run a macro with a VBA error trap so runtime errors can be returned to Electron.
   * @param {string} macroName
   * @returns {{ success: boolean, message: string, error?: { number: number, description: string, source: string } }}
   */
  runMacroWithTrap(macroName) {
    try {
      this._ensureRuntimeModule();
      const excel = this.getApp();
      const workbook = this.getActiveWorkbook();
      const workbookName = this._qualifyWorkbookName(workbook.Name);
      const runtimeMacro = `${workbookName}!MacroFlow_Runtime.MacroFlow_RunMacro`;
      const targetMacro = this._normalizeMacroName(macroName, workbook);

      if (!targetMacro) {
        return { success: false, message: 'Macro name is required.' };
      }

      const prep = this._prepareMacroRun(excel, workbook);
      if (!prep.ready) {
        return { success: false, message: prep.message };
      }

      const result = excel.Run(runtimeMacro, targetMacro);

      if (typeof result === 'string' && result.startsWith('ERR|')) {
        const [, number, description, source] = result.split('|');
        return {
          success: false,
          message: `VBA error ${number}: ${description}`,
          error: {
            number: Number(number),
            description: description || '',
            source: source || ''
          }
        };
      }

      return { success: true, message: `Executed "${targetMacro}"` };
    } catch (error) {
      const rawMessage = error && error.message ? String(error.message) : 'Unknown error';
      const hresultMatch = rawMessage.match(/0x[0-9a-fA-F]+/);
      const hresult = hresultMatch ? hresultMatch[0] : '';
      const isExcelRunFailure = rawMessage.includes('0x800a9c68');
      const message = isExcelRunFailure
        ? 'Excel failed before VBA executed (likely compile error, missing reference, or dialog).'
        : `Failed to run macro: ${rawMessage}`;

      return {
        success: false,
        message,
        error: {
          kind: isExcelRunFailure ? 'excel-run-failed' : 'runtime',
          hresult,
          rawMessage
        }
      };
    }
  }

  _loadShortcutRegistry(workbook) {
    const props = workbook.CustomDocumentProperties;
    const registryName = 'MacroFlow_Shortcuts';
    let prop = null;

    try {
      prop = props.Item(registryName);
    } catch (error) {
      prop = null;
    }

    if (!prop) {
      props.Add(registryName, false, 4, '{}');
      prop = props.Item(registryName);
    }

    try {
      return JSON.parse(String(prop.Value || '{}'));
    } catch (error) {
      return {};
    }
  }

  _saveShortcutRegistry(workbook, registry) {
    const props = workbook.CustomDocumentProperties;
    const registryName = 'MacroFlow_Shortcuts';
    let prop = null;

    try {
      prop = props.Item(registryName);
    } catch (error) {
      prop = null;
    }

    if (!prop) {
      props.Add(registryName, false, 4, JSON.stringify(registry));
      return;
    }

    prop.Value = JSON.stringify(registry);
  }

  _listMacroProceduresForWorkbook(workbook) {
    try {
      const vbProject = this._getVBProjectForWorkbook(workbook);
      const procedures = this._listProceduresForWorkbook(workbook, vbProject);
      return procedures
        .filter((proc) => String(proc?.kind || '').startsWith('Sub'))
        .map((proc) => ({
          name: proc.name,
          module: proc.module
        }));
    } catch (error) {
      this._proceduresCache = null;
      return [];
    }
  }

  _listMacroProcedures(options = {}) {
    const { activate = true } = options;
    try {
      const workbook = this.getActiveWorkbook({ activate });
      const procedures = this._listMacroProceduresForWorkbook(workbook);
      return { workbook, procedures };
    } catch (error) {
      this._proceduresCache = null;
      return { workbook: null, procedures: [] };
    }
  }

  _setMacroShortcutForWorkbook(workbook, macroName, shortcutKey) {
    const excel = this.getApp();
    const procedures = this._listMacroProceduresForWorkbook(workbook);
    const resolvedName = this._normalizeMacroNameForWorkbook(macroName, workbook, procedures);
    excel.MacroOptions(resolvedName || macroName, null, null, null, null, shortcutKey, null, null, null, null);

    const registry = this._loadShortcutRegistry(workbook);
    if (macroName) {
      registry[macroName] = shortcutKey;
    }
    if (resolvedName) {
      registry[resolvedName] = shortcutKey;
    }
    this._saveShortcutRegistry(workbook, registry);

    return { success: true, message: `Shortcut set for ${macroName}` };
  }

  /**
   * Set a macro shortcut and track it in a workbook registry.
   * @param {string} macroName
   * @param {string} shortcutKey
   * @returns {{ success: boolean, message: string }}
   */
  setMacroShortcut(macroName, shortcutKey) {
    try {
      const workbook = this.getActiveWorkbook();
      return this._setMacroShortcutForWorkbook(workbook, macroName, shortcutKey);
    } catch (error) {
      return { success: false, message: `Failed to set shortcut: ${error.message}` };
    }
  }

  /**
   * Set a macro shortcut in a specific open workbook and track it in that workbook registry.
   * @param {string} workbookName
   * @param {string} macroName
   * @param {string} shortcutKey
   * @param {{ activate?: boolean }} options
   * @returns {{ success: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, message: string }}
   */
  setMacroShortcutByWorkbookName(workbookName, macroName, shortcutKey, options = {}) {
    const { activate = true } = options;
    const normalizedName = String(workbookName || '').trim();
    if (!normalizedName) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        message: 'Workbook name is required.'
      };
    }

    try {
      const excel = this.getApp({ activate });
      const workbook = this._findOpenWorkbookByName(excel, normalizedName);
      if (!workbook) {
        return {
          success: true,
          workbookFound: false,
          workbook: null,
          message: `Workbook "${normalizedName}" is not open.`
        };
      }

      const result = this._setMacroShortcutForWorkbook(workbook, macroName, shortcutKey);
      return {
        ...result,
        workbookFound: true,
        workbook: this._describeWorkbook(workbook)
      };
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        message: `Failed to set shortcut: ${error.message}`
      };
    }
  }

  _auditShortcutsForWorkbook(workbook) {
    const procedures = this._listMacroProceduresForWorkbook(workbook);
    const registry = this._loadShortcutRegistry(workbook);
    const shortcuts = [];
    const unmapped = [];

    procedures.forEach((proc) => {
      const fullName = `${proc.module}.${proc.name}`;
      const qualifiedName = `${workbook.Name}!${fullName}`;
      const shortcut =
        registry[qualifiedName] ||
        registry[fullName] ||
        registry[proc.name] ||
        null;
      if (shortcut) {
        shortcuts.push({
          macro: qualifiedName,
          shortcut
        });
      } else {
        unmapped.push(qualifiedName);
      }
    });

    return {
      success: true,
      shortcuts,
      unmapped,
      note: 'Excel does not expose global shortcut listings. Only MacroFlow-tracked shortcuts are available.'
    };
  }

  /**
   * Audit known shortcuts for macros in the active workbook.
   * Note: Excel does not expose a full shortcut map, so only tracked shortcuts are returned.
   * @returns {{ success: boolean, shortcuts: Array, unmapped: Array, note: string }}
   */
  auditShortcuts() {
    try {
      const workbook = this.getActiveWorkbook();
      if (!workbook) {
        return { success: false, shortcuts: [], unmapped: [], message: 'Workbook not available.' };
      }
      return this._auditShortcutsForWorkbook(workbook);
    } catch (error) {
      return { success: false, shortcuts: [], unmapped: [], message: error.message };
    }
  }

  /**
   * Audit known shortcuts for macros in a specific open workbook.
   * @param {string} workbookName
   * @param {{ activate?: boolean }} options
   * @returns {{ success: boolean, workbookFound: boolean, workbook?: { name: string, path: string } | null, shortcuts: Array, unmapped: Array, note?: string, message?: string }}
   */
  auditShortcutsByWorkbookName(workbookName, options = {}) {
    const { activate = true } = options;
    const normalizedName = String(workbookName || '').trim();
    if (!normalizedName) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        shortcuts: [],
        unmapped: [],
        message: 'Workbook name is required.'
      };
    }

    try {
      const excel = this.getApp({ activate });
      const workbook = this._findOpenWorkbookByName(excel, normalizedName);
      if (!workbook) {
        return {
          success: true,
          workbookFound: false,
          workbook: null,
          shortcuts: [],
          unmapped: [],
          message: `Workbook "${normalizedName}" is not open.`
        };
      }

      const result = this._auditShortcutsForWorkbook(workbook);
      return {
        ...result,
        workbookFound: true,
        workbook: this._describeWorkbook(workbook)
      };
    } catch (error) {
      return {
        success: false,
        workbookFound: false,
        workbook: null,
        shortcuts: [],
        unmapped: [],
        message: error.message
      };
    }
  }
}

module.exports = new ExcelBridge();
