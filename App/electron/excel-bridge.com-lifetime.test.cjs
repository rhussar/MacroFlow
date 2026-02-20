const assert = require('node:assert/strict');
const test = require('node:test');
const path = require('node:path');
const Module = require('node:module');

const BRIDGE_PATH = path.resolve(__dirname, 'excel-bridge.js');

function buildTasklistOutput(processIds = []) {
  if (!Array.isArray(processIds) || processIds.length < 1) {
    return 'INFO: No tasks are running which match the specified criteria.\r\n';
  }

  return processIds
    .map((pid) => `"EXCEL.EXE","${pid}","Console","1","123,456 K"`)
    .join('\r\n');
}

function loadExcelBridge({
  processIds = [4242],
  objectFactory = () => ({})
} = {}) {
  const originalLoad = Module._load;
  const releaseCalls = [];
  let objectCalls = 0;

  const winaxStub = {
    Object: function ObjectFactory(id, options) {
      objectCalls += 1;
      return objectFactory({ id, options, call: objectCalls });
    },
    release: (...objects) => {
      releaseCalls.push(objects);
    }
  };

  const childProcessStub = {
    execSync: () => buildTasklistOutput(processIds)
  };

  Module._load = function patchedLoad(request, parent, isMain) {
    if (request === 'winax') {
      return winaxStub;
    }
    if (request === 'node:child_process') {
      return childProcessStub;
    }
    return originalLoad.call(this, request, parent, isMain);
  };

  delete require.cache[BRIDGE_PATH];
  const bridge = require(BRIDGE_PATH);
  Module._load = originalLoad;

  return {
    bridge,
    releaseCalls,
    getObjectCalls: () => objectCalls
  };
}

test('core operations release COM handles on completion', () => {
  const workbook = {
    Name: 'Book1.xlsx',
    FullName: 'C:\\Book1.xlsx',
    ActiveSheet: { Name: 'Sheet1' },
    Sheets: {
      Count: 1,
      Item: () => ({ Name: 'Sheet1' })
    }
  };

  const range = { Value2: 42 };
  const excelApp = {
    ActiveWorkbook: workbook,
    ActiveSheet: {
      Range: () => range
    },
    Workbooks: {
      Count: 1,
      Item: () => workbook
    }
  };

  const { bridge, releaseCalls } = loadExcelBridge({
    processIds: [1111],
    objectFactory: () => excelApp
  });

  const beforeInfo = releaseCalls.length;
  const info = bridge.getWorkbookInfo();
  assert.equal(info.success, true);
  assert.ok(releaseCalls.length > beforeInfo, 'getWorkbookInfo should release COM handles');

  const beforeList = releaseCalls.length;
  const list = bridge.getOpenWorkbooks();
  assert.equal(list.success, true);
  assert.ok(releaseCalls.length > beforeList, 'getOpenWorkbooks should release COM handles');

  const beforeRead = releaseCalls.length;
  const read = bridge.readCell('A1');
  assert.equal(read.success, true);
  assert.ok(releaseCalls.length > beforeRead, 'readCell should release COM handles');
});

test('failure paths release COM handles', () => {
  const excelApp = { ActiveWorkbook: null };
  const { bridge, releaseCalls } = loadExcelBridge({
    processIds: [2222],
    objectFactory: () => excelApp
  });

  const result = bridge.getWorkbookInfo();
  assert.equal(result.success, false);
  assert.match(result.message, /NO_WORKBOOK/);

  const releasedObjects = releaseCalls.flat();
  assert.ok(releasedObjects.includes(excelApp), 'Excel application should be released on failure');
});

test('NO_EXCEL preflight occurs before COM attach', () => {
  const { bridge, getObjectCalls } = loadExcelBridge({
    processIds: [],
    objectFactory: () => {
      throw new Error('COM attach should not be attempted');
    }
  });

  const result = bridge.getWorkbookInfo();
  assert.equal(result.success, false);
  assert.match(result.message, /NO_EXCEL/);
  assert.equal(getObjectCalls(), 0);
});

test('getOpenWorkbooks uses fresh COM attach for each call (no persistent cache)', () => {
  let nextPid = 3000;
  const { bridge, getObjectCalls } = loadExcelBridge({
    processIds: [3333],
    objectFactory: () => ({
      id: nextPid++,
      Workbooks: {
        Count: 0
      }
    })
  });

  const first = bridge.getOpenWorkbooks();
  const second = bridge.getOpenWorkbooks();

  assert.equal(first.success, true);
  assert.equal(second.success, true);
  assert.equal(getObjectCalls(), 2);
});

test('auditShortcuts releases workbook property COM handles', () => {
  const shortcutProp = { Value: '{}' };
  const customProps = {
    Item: () => shortcutProp,
    Add: () => {}
  };
  const vbComponents = { Count: 0 };
  const vbProject = { VBComponents: vbComponents };
  const workbook = {
    Name: 'Book1.xlsx',
    FullName: 'C:\\Book1.xlsx',
    VBProject: vbProject,
    CustomDocumentProperties: customProps
  };
  const excelApp = {
    ActiveWorkbook: workbook
  };

  const { bridge, releaseCalls } = loadExcelBridge({
    processIds: [4444],
    objectFactory: () => excelApp
  });

  const result = bridge.auditShortcuts();
  assert.equal(result.success, true);

  const releasedObjects = releaseCalls.flat();
  assert.ok(releasedObjects.includes(customProps), 'CustomDocumentProperties should be released');
  assert.ok(releasedObjects.includes(shortcutProp), 'Shortcut registry property should be released');
});

test('shutdown latch blocks COM attach immediately', () => {
  const { bridge, getObjectCalls } = loadExcelBridge({
    processIds: [5555],
    objectFactory: () => {
      throw new Error('COM attach should not be attempted during shutdown');
    }
  });

  bridge.setShuttingDown(true);
  const result = bridge.getWorkbookInfo();
  assert.equal(result.success, false);
  assert.match(result.message, /APP_SHUTTING_DOWN/);
  assert.equal(getObjectCalls(), 0);
});

test('getActiveWorkbookContext releases COM handles and returns combined payload', () => {
  const shortcutProp = { Value: '{}' };
  const customProps = {
    Item: () => shortcutProp,
    Add: () => {}
  };
  const vbComponents = { Count: 0 };
  const vbProject = { VBComponents: vbComponents };
  const workbook = {
    Name: 'Book1.xlsx',
    FullName: 'C:\\Book1.xlsx',
    ActiveSheet: { Name: 'Sheet1' },
    Sheets: {
      Count: 1,
      Item: () => ({ Name: 'Sheet1' })
    },
    VBProject: vbProject,
    CustomDocumentProperties: customProps
  };
  const excelApp = { ActiveWorkbook: workbook };

  const { bridge, releaseCalls } = loadExcelBridge({
    processIds: [6666],
    objectFactory: () => excelApp
  });

  const result = bridge.getActiveWorkbookContext();
  assert.equal(result.success, true);
  assert.equal(result.workbook?.name, 'Book1.xlsx');
  assert.ok(Array.isArray(result.modules));
  assert.ok(Array.isArray(result.procedures));
  assert.equal(result.shortcutAudit?.success, true);

  const releasedObjects = releaseCalls.flat();
  assert.ok(releasedObjects.includes(excelApp), 'Excel application should be released');
  assert.ok(releasedObjects.includes(workbook), 'Active workbook should be released');
  assert.ok(releasedObjects.includes(customProps), 'CustomDocumentProperties should be released');
});

test('getOpenWorkbookListContext releases COM handles and returns list payload', () => {
  const vbProject = { VBComponents: { Count: 0 } };
  const workbook = {
    Name: 'Book1.xlsx',
    FullName: 'C:\\Book1.xlsx',
    VBProject: vbProject
  };
  const workbooksCollection = {
    Count: 1,
    Item: () => workbook
  };
  const excelApp = {
    Workbooks: workbooksCollection
  };

  const { bridge, releaseCalls } = loadExcelBridge({
    processIds: [7777],
    objectFactory: () => excelApp
  });

  const result = bridge.getOpenWorkbookListContext();
  assert.equal(result.success, true);
  assert.equal(result.workbooks.length, 1);
  assert.ok(Array.isArray(result.allFilesModules));

  const releasedObjects = releaseCalls.flat();
  assert.ok(releasedObjects.includes(excelApp), 'Excel application should be released');
  assert.ok(releasedObjects.includes(workbook), 'Workbook should be released');
  assert.ok(releasedObjects.includes(workbooksCollection), 'Workbook collection should be released');
});

test('injectModuleByWorkbookName creates a module in the requested workbook', () => {
  let insertedLine = 0;
  let insertedCode = '';
  const codeModule = {
    CountOfLines: 0,
    InsertLines: (line, code) => {
      insertedLine = line;
      insertedCode = code;
    }
  };
  const newModule = {
    Name: '',
    CodeModule: codeModule
  };
  const existingComponent = {
    Name: 'OtherModule',
    CodeModule: { CountOfLines: 0 }
  };
  const vbComponents = {
    Count: 1,
    Item: () => existingComponent,
    Add: () => newModule,
    Remove: () => {}
  };
  const vbProject = { VBComponents: vbComponents };
  const targetWorkbook = {
    Name: 'Client.xlsm',
    FullName: 'C:\\Client.xlsm',
    VBProject: vbProject
  };
  const otherWorkbook = {
    Name: 'Other.xlsm',
    FullName: 'C:\\Other.xlsm'
  };
  const excelApp = {
    Workbooks: {
      Count: 2,
      Item: (index) => (index === 1 ? otherWorkbook : targetWorkbook)
    }
  };

  const { bridge } = loadExcelBridge({
    processIds: [8888],
    objectFactory: () => excelApp
  });

  const result = bridge.injectModuleByWorkbookName(
    'Client.xlsm',
    'MacroFlowModule1',
    'Sub RunA()\nEnd Sub',
    { workbookPath: 'C:\\Client.xlsm' }
  );

  assert.equal(result.success, true);
  assert.equal(result.workbookFound, true);
  assert.equal(result.workbook?.name, 'Client.xlsm');
  assert.equal(result.moduleName, 'MacroFlowModule1');
  assert.equal(newModule.Name, 'MacroFlowModule1');
  assert.equal(insertedLine, 1);
  assert.equal(insertedCode, 'Sub RunA()\nEnd Sub');
});

test('injectModuleByWorkbookName returns workbookFound false when workbook is missing', () => {
  const workbook = {
    Name: 'Open.xlsm',
    FullName: 'C:\\Open.xlsm'
  };
  const excelApp = {
    Workbooks: {
      Count: 1,
      Item: () => workbook
    }
  };

  const { bridge } = loadExcelBridge({
    processIds: [9999],
    objectFactory: () => excelApp
  });

  const result = bridge.injectModuleByWorkbookName(
    'Missing.xlsm',
    'MacroFlowModule1',
    'Sub RunA()\nEnd Sub',
    { workbookPath: 'C:\\Missing.xlsm' }
  );

  assert.equal(result.success, true);
  assert.equal(result.workbookFound, false);
  assert.equal(result.workbook, null);
});

test('moduleCodeByWorkbookName returns code, lineCount, and hash for target module', () => {
  const codeModule = {
    CountOfLines: 3,
    Lines: () => 'Option Explicit\r\nSub RunA()\r\nEnd Sub'
  };
  const component = {
    Name: 'MacroFlowModule1',
    CodeModule: codeModule
  };
  const vbComponents = {
    Count: 1,
    Item: () => component
  };
  const vbProject = { VBComponents: vbComponents };
  const workbook = {
    Name: 'Client.xlsm',
    FullName: 'C:\\Client.xlsm',
    VBProject: vbProject
  };
  const excelApp = {
    Workbooks: {
      Count: 1,
      Item: () => workbook
    }
  };

  const { bridge } = loadExcelBridge({
    processIds: [10001],
    objectFactory: () => excelApp
  });

  const result = bridge.getModuleCodeByWorkbookName(
    'Client.xlsm',
    'MacroFlowModule1',
    { workbookPath: 'C:\\Client.xlsm' }
  );

  assert.equal(result.success, true);
  assert.equal(result.workbookFound, true);
  assert.equal(result.moduleFound, true);
  assert.equal(result.moduleName, 'MacroFlowModule1');
  assert.equal(result.lineCount, 3);
  assert.equal(result.code, 'Option Explicit\nSub RunA()\nEnd Sub');
  assert.equal(result.hash, '8b823a0a');
});

test('moduleSignatureByWorkbookName omits code and reports missing module/workbook surfaces', () => {
  const vbComponents = { Count: 0 };
  const vbProject = { VBComponents: vbComponents };
  const workbook = {
    Name: 'Client.xlsm',
    FullName: 'C:\\Client.xlsm',
    VBProject: vbProject
  };
  const excelApp = {
    Workbooks: {
      Count: 1,
      Item: () => workbook
    }
  };

  const { bridge } = loadExcelBridge({
    processIds: [10002],
    objectFactory: () => excelApp
  });

  const missingModuleResult = bridge.getModuleSignatureByWorkbookName(
    'Client.xlsm',
    'MissingModule',
    { workbookPath: 'C:\\Client.xlsm' }
  );
  assert.equal(missingModuleResult.success, true);
  assert.equal(missingModuleResult.workbookFound, true);
  assert.equal(missingModuleResult.moduleFound, false);
  assert.equal(Object.prototype.hasOwnProperty.call(missingModuleResult, 'code'), false);

  const missingWorkbookResult = bridge.getModuleSignatureByWorkbookName(
    'Missing.xlsm',
    'MissingModule',
    { workbookPath: 'C:\\Missing.xlsm' }
  );
  assert.equal(missingWorkbookResult.success, true);
  assert.equal(missingWorkbookResult.workbookFound, false);
  assert.equal(missingWorkbookResult.moduleFound, false);
});

test('setModuleCodeByWorkbookName updates existing module and can create when missing', () => {
  const existingCodeState = { text: 'Sub Old()\nEnd Sub' };
  const existingCodeModule = {
    get CountOfLines() {
      return existingCodeState.text ? existingCodeState.text.split('\n').length : 0;
    },
    Lines: () => existingCodeState.text,
    DeleteLines: () => {
      existingCodeState.text = '';
    },
    InsertLines: (_line, text) => {
      existingCodeState.text = String(text || '');
    }
  };
  const existingComponent = {
    Name: 'MacroFlowModule1',
    CodeModule: existingCodeModule
  };

  const createdCodeState = { text: '' };
  const createdCodeModule = {
    get CountOfLines() {
      return createdCodeState.text ? createdCodeState.text.split('\n').length : 0;
    },
    Lines: () => createdCodeState.text,
    DeleteLines: () => {
      createdCodeState.text = '';
    },
    InsertLines: (_line, text) => {
      createdCodeState.text = String(text || '');
    }
  };

  const components = [existingComponent];
  const vbComponents = {
    get Count() {
      return components.length;
    },
    Item: (index) => components[index - 1],
    Add: () => {
      const createdComponent = { Name: '', CodeModule: createdCodeModule };
      components.push(createdComponent);
      return createdComponent;
    },
    Remove: () => {}
  };
  const vbProject = { VBComponents: vbComponents };
  const workbook = {
    Name: 'Client.xlsm',
    FullName: 'C:\\Client.xlsm',
    VBProject: vbProject
  };
  const excelApp = {
    Workbooks: {
      Count: 1,
      Item: () => workbook
    }
  };

  const { bridge } = loadExcelBridge({
    processIds: [10003],
    objectFactory: () => excelApp
  });

  const updated = bridge.setModuleCodeByWorkbookName(
    'Client.xlsm',
    'MacroFlowModule1',
    'Option Explicit\nSub Updated()\nEnd Sub',
    { workbookPath: 'C:\\Client.xlsm', createIfMissing: false }
  );
  assert.equal(updated.success, true);
  assert.equal(updated.workbookFound, true);
  assert.equal(updated.moduleFound, true);
  assert.equal(updated.moduleName, 'MacroFlowModule1');
  assert.equal(existingCodeState.text, 'Option Explicit\nSub Updated()\nEnd Sub');
  assert.equal(updated.lineCount, 3);
  assert.equal(updated.hash, '2bae8529');

  const missingWithoutCreate = bridge.setModuleCodeByWorkbookName(
    'Client.xlsm',
    'MacroFlowModule2',
    'Option Explicit',
    { workbookPath: 'C:\\Client.xlsm', createIfMissing: false }
  );
  assert.equal(missingWithoutCreate.success, false);
  assert.equal(missingWithoutCreate.workbookFound, true);
  assert.equal(missingWithoutCreate.moduleFound, false);

  const created = bridge.setModuleCodeByWorkbookName(
    'Client.xlsm',
    'MacroFlowModule2',
    'Option Explicit',
    { workbookPath: 'C:\\Client.xlsm', createIfMissing: true }
  );
  assert.equal(created.success, true);
  assert.equal(created.workbookFound, true);
  assert.equal(created.moduleFound, false);
  assert.equal(created.moduleName, 'MacroFlowModule2');
  assert.equal(components.length, 2);
  assert.equal(components[1].Name, 'MacroFlowModule2');
  assert.equal(createdCodeState.text, 'Option Explicit');
  assert.equal(created.hash, '621fb430');
});

test('module code operations release COM handles for workbook/code module paths', () => {
  const codeModule = {
    CountOfLines: 1,
    Lines: () => 'Option Explicit'
  };
  const component = {
    Name: 'MacroFlowModule1',
    CodeModule: codeModule
  };
  const vbComponents = {
    Count: 1,
    Item: () => component,
    Add: () => component
  };
  const vbProject = { VBComponents: vbComponents };
  const workbook = {
    Name: 'Client.xlsm',
    FullName: 'C:\\Client.xlsm',
    VBProject: vbProject
  };
  const workbooksCollection = {
    Count: 1,
    Item: () => workbook
  };
  const excelApp = {
    Workbooks: workbooksCollection
  };

  const { bridge, releaseCalls } = loadExcelBridge({
    processIds: [10004],
    objectFactory: () => excelApp
  });

  const before = releaseCalls.length;
  const result = bridge.getModuleCodeByWorkbookName(
    'Client.xlsm',
    'MacroFlowModule1',
    { workbookPath: 'C:\\Client.xlsm' }
  );
  assert.equal(result.success, true);
  assert.ok(releaseCalls.length > before, 'module code read should release COM handles');

  const released = releaseCalls.flat();
  assert.ok(released.includes(excelApp), 'Excel app should be released');
  assert.ok(released.includes(workbook), 'Workbook should be released');
  assert.ok(released.includes(vbProject), 'VBProject should be released');
  assert.ok(released.includes(codeModule), 'CodeModule should be released');
});

test('renameModuleByWorkbookName renames only standard modules and blocks duplicates/invalid names', () => {
  const standardComponent = {
    Name: 'ModuleOne',
    Type: 1
  };
  const existingComponent = {
    Name: 'ModuleTwo',
    Type: 1
  };
  const classComponent = {
    Name: 'ClassOne',
    Type: 2
  };
  const components = [standardComponent, existingComponent, classComponent];
  const vbComponents = {
    get Count() {
      return components.length;
    },
    Item: (index) => components[index - 1],
    Remove: () => {}
  };
  const vbProject = { VBComponents: vbComponents };
  const workbook = {
    Name: 'Client.xlsm',
    FullName: 'C:\\Client.xlsm',
    VBProject: vbProject
  };
  const excelApp = {
    Workbooks: {
      Count: 1,
      Item: () => workbook
    }
  };

  const { bridge } = loadExcelBridge({
    processIds: [11001],
    objectFactory: () => excelApp
  });

  const renamed = bridge.renameModuleByWorkbookName(
    'Client.xlsm',
    'ModuleOne',
    'RenamedModule',
    { workbookPath: 'C:\\Client.xlsm' }
  );
  assert.equal(renamed.success, true);
  assert.equal(renamed.workbookFound, true);
  assert.equal(renamed.moduleFound, true);
  assert.equal(renamed.renamed, true);
  assert.equal(standardComponent.Name, 'RenamedModule');

  const duplicate = bridge.renameModuleByWorkbookName(
    'Client.xlsm',
    'RenamedModule',
    'ModuleTwo',
    { workbookPath: 'C:\\Client.xlsm' }
  );
  assert.equal(duplicate.success, false);
  assert.equal(duplicate.renamed, false);

  const invalidName = bridge.renameModuleByWorkbookName(
    'Client.xlsm',
    'RenamedModule',
    '1-invalid',
    { workbookPath: 'C:\\Client.xlsm' }
  );
  assert.equal(invalidName.success, false);
  assert.equal(invalidName.renamed, false);

  const nonStandard = bridge.renameModuleByWorkbookName(
    'Client.xlsm',
    'ClassOne',
    'ClassRenamed',
    { workbookPath: 'C:\\Client.xlsm' }
  );
  assert.equal(nonStandard.success, false);
  assert.equal(nonStandard.moduleFound, true);
  assert.equal(nonStandard.renamed, false);
});

test('deleteModuleByWorkbookName deletes standard modules and blocks non-standard modules', () => {
  const standardComponent = {
    Name: 'DeleteMe',
    Type: 1
  };
  const classComponent = {
    Name: 'ClassKeep',
    Type: 2
  };
  const components = [standardComponent, classComponent];
  const vbComponents = {
    get Count() {
      return components.length;
    },
    Item: (index) => components[index - 1],
    Remove: (component) => {
      const index = components.indexOf(component);
      if (index >= 0) {
        components.splice(index, 1);
      }
    }
  };
  const vbProject = { VBComponents: vbComponents };
  const workbook = {
    Name: 'Client.xlsm',
    FullName: 'C:\\Client.xlsm',
    VBProject: vbProject
  };
  const excelApp = {
    Workbooks: {
      Count: 1,
      Item: () => workbook
    }
  };

  const { bridge } = loadExcelBridge({
    processIds: [11002],
    objectFactory: () => excelApp
  });

  const deleted = bridge.deleteModuleByWorkbookName(
    'Client.xlsm',
    'DeleteMe',
    { workbookPath: 'C:\\Client.xlsm' }
  );
  assert.equal(deleted.success, true);
  assert.equal(deleted.deleted, true);
  assert.equal(components.some((component) => component.Name === 'DeleteMe'), false);

  const nonStandard = bridge.deleteModuleByWorkbookName(
    'Client.xlsm',
    'ClassKeep',
    { workbookPath: 'C:\\Client.xlsm' }
  );
  assert.equal(nonStandard.success, false);
  assert.equal(nonStandard.deleted, false);
  assert.equal(nonStandard.moduleFound, true);
});

test('module rename/delete operations release COM handles', () => {
  const standardComponent = {
    Name: 'ModuleOne',
    Type: 1
  };
  const components = [standardComponent];
  const vbComponents = {
    get Count() {
      return components.length;
    },
    Item: (index) => components[index - 1],
    Remove: () => {}
  };
  const vbProject = { VBComponents: vbComponents };
  const workbook = {
    Name: 'Client.xlsm',
    FullName: 'C:\\Client.xlsm',
    VBProject: vbProject
  };
  const excelApp = {
    Workbooks: {
      Count: 1,
      Item: () => workbook
    }
  };

  const { bridge, releaseCalls } = loadExcelBridge({
    processIds: [11003],
    objectFactory: () => excelApp
  });

  const renameBefore = releaseCalls.length;
  const renameResult = bridge.renameModuleByWorkbookName(
    'Client.xlsm',
    'ModuleOne',
    'ModuleRenamed',
    { workbookPath: 'C:\\Client.xlsm' }
  );
  assert.equal(renameResult.success, true);
  assert.ok(releaseCalls.length > renameBefore);

  const deleteBefore = releaseCalls.length;
  const deleteResult = bridge.deleteModuleByWorkbookName(
    'Client.xlsm',
    'ModuleRenamed',
    { workbookPath: 'C:\\Client.xlsm' }
  );
  assert.equal(deleteResult.success, true);
  assert.ok(releaseCalls.length > deleteBefore);

  const released = releaseCalls.flat();
  assert.ok(released.includes(excelApp), 'Excel app should be released');
  assert.ok(released.includes(workbook), 'Workbook should be released');
  assert.ok(released.includes(vbProject), 'VBProject should be released');
});
