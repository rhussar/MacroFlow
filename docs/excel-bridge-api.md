# Excel Bridge API (Current)

This document lists the APIs exposed to the renderer via `window.excel` in
`App/electron/preload.js`. These are the **current** calls that exist today.

## Usage

```js
const result = await window.excel.workbook.info();
```

All calls return a Promise that resolves to an object with at least:

- `success: boolean`
- `message?: string` (when provided by the handler)

## App Controls

### `app.close()`
Closes the application.

---

## VBA Operations

### `vba.inject({ moduleName?: string, code: string })`
Inject VBA code into a module.

Returns: `{ success: boolean, message: string }`

### `vba.injectByWorkbook({ workbookName?: string, workbookPath?: string, moduleName: string, code: string, createIfMissing?: boolean })`
Inject VBA code into a module in a specific open workbook.

Returns: `{ success: boolean, workbookFound: boolean, workbook?: object, moduleName?: string, message: string }`

### `vba.moduleCodeByWorkbook({ workbookName?: string, workbookPath?: string, moduleName: string })`
Read module code in a specific open workbook.

Returns: `{ success: boolean, workbookFound: boolean, moduleFound: boolean, workbook?: object, moduleName?: string, lineCount?: number, hash?: string, code?: string, message?: string }`

### `vba.moduleSignatureByWorkbook({ workbookName?: string, workbookPath?: string, moduleName: string })`
Read module signature (line count + hash) in a specific open workbook.

Returns: `{ success: boolean, workbookFound: boolean, moduleFound: boolean, workbook?: object, moduleName?: string, lineCount?: number, hash?: string, message?: string }`

### `vba.setModuleCodeByWorkbook({ workbookName?: string, workbookPath?: string, moduleName: string, code: string, createIfMissing?: boolean })`
Set module code in a specific open workbook.

Returns: `{ success: boolean, workbookFound: boolean, moduleFound: boolean, workbook?: object, moduleName?: string, lineCount?: number, hash?: string, message?: string }`

### `vba.run({ macroName: string })`
Run a VBA macro.

Returns: `{ success: boolean, message: string }`

### `vba.modules()`
List VBA modules in the active workbook.

Returns: `{ success: boolean, workbook?: object, modules: Array }`

### `vba.procedures()`
List procedures (Subs/Functions/Properties) in the active workbook.

Returns: `{ success: boolean, workbook?: object, procedures: Array }`

### `vba.setShortcut({ macroName: string, shortcutKey: string })`
Set a macro shortcut and track it.

Returns: `{ success: boolean, message: string }`

### `vba.auditShortcuts()`
Audit tracked shortcuts.

Returns: `{ success: boolean, shortcuts: Array, unmapped: Array }`

---

## Cell Operations

### `cell.read({ address: string })`
Read a cell value.

Returns: `{ success: boolean, address: string, value: any }`

### `cell.write({ address: string, value: any })`
Write a value to a cell.

Returns: `{ success: boolean, address: string }`

### `cell.selection()`
Get the current selection.

Returns: `{ success: boolean, address: string, value: any }`

### `cell.highlight({ color: string })`
Highlight current selection.

Returns: `{ success: boolean, color?: string }`

---

## Workbook Operations

### `workbook.info()`
Get info about the active workbook.

Returns: `{ success: boolean, name: string, path: string, sheets: string[] }`

### `workbook.list()`
Get all open workbooks.

Returns: `{ success: boolean, workbooks: Array<{ name: string, path: string }> }`

### `workbook.sheets()`
List worksheets with UsedRange stats.

Returns: `{ success: boolean, sheets: Array }`

### `workbook.metadata({ sheetName?: string })`
Get worksheet metadata and preview.

Returns: `{ success: boolean, sheet?: object, structuralContext?: object, dataContext?: object }`

### `workbook.metadataClosed({ path: string, sheetName?: string })`
Get metadata for a closed workbook path.

Returns: `{ success: boolean, sheet?: object, structuralContext?: object, dataContext?: object }`
