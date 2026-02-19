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
