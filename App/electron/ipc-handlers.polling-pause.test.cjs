const assert = require('node:assert/strict');
const test = require('node:test');
const path = require('node:path');
const Module = require('node:module');

const IPC_HANDLERS_PATH = path.resolve(__dirname, 'ipc-handlers.js');

function loadHandlers({ excelOverrides = {} } = {}) {
  const originalLoad = Module._load;
  const handlers = {};
  let clearComCacheCalls = 0;

  const excelStub = {
    clearComCache: () => {
      clearComCacheCalls += 1;
    },
    getWorkbookInfo: () => ({ success: true, name: 'Book1' }),
    listModules: () => ({ success: true, modules: [] }),
    listProcedures: () => ({ success: true, procedures: [] }),
    getOpenWorkbooks: () => ({ success: true, workbooks: [] }),
    auditShortcuts: () => ({ success: true, shortcuts: [], unmapped: [] }),
    auditShortcutsByWorkbookName: () => ({ success: true, shortcuts: [], unmapped: [], workbookFound: true }),
    getSelection: () => ({ success: true, address: 'A1', value: 'x' }),
    ...excelOverrides
  };

  const electronStub = {
    ipcMain: {
      handle: (channel, handler) => {
        handlers[channel] = handler;
      },
      on: () => {}
    },
    app: {
      quit: () => {}
    },
    BrowserWindow: {
      getAllWindows: () => []
    }
  };

  const diagnosticsStub = {
    logger: {
      debug: () => {},
      warn: () => {}
    },
    checkExcelModalState: async () => ({ hasModal: false }),
    collectDiagnostics: async () => ({}),
    checkAddinStatus: async () => ({}),
    checkRibbonStatus: async () => ({})
  };

  Module._load = function patchedLoad(request, parent, isMain) {
    if (request === 'electron') {
      return electronStub;
    }
    if (request === './excel-bridge') {
      return excelStub;
    }
    if (request === './diagnostics') {
      return diagnosticsStub;
    }
    return originalLoad.call(this, request, parent, isMain);
  };

  delete require.cache[IPC_HANDLERS_PATH];
  const ipcHandlers = require(IPC_HANDLERS_PATH);
  ipcHandlers.registerHandlers();
  Module._load = originalLoad;

  return {
    handlers,
    excelStub,
    getClearComCacheCalls: () => clearComCacheCalls
  };
}

test('polling pause latches after NO_EXCEL and protected channels short-circuit', async () => {
  let getWorkbookInfoCalls = 0;
  let listModulesCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => {
        getWorkbookInfoCalls += 1;
        return { success: false, message: 'NO_EXCEL: Excel is not running.' };
      },
      listModules: () => {
        listModulesCalls += 1;
        return { success: true, modules: [{ name: 'Module1' }] };
      }
    }
  });

  const workbookInfoResult = await handlers['workbook:info']();
  assert.equal(workbookInfoResult.success, false);
  assert.match(workbookInfoResult.message, /NO_EXCEL/);
  assert.equal(getWorkbookInfoCalls, 1);

  const modulesResult = await handlers['vba:modules']();
  assert.equal(modulesResult.success, false);
  assert.equal(modulesResult.paused, true);
  assert.equal(modulesResult.reason, 'polling_paused');
  assert.match(modulesResult.message, /NO_EXCEL/);
  assert.deepEqual(modulesResult.modules, []);
  assert.equal(listModulesCalls, 0);
});

test('excel:reconnect success clears polling pause and protected channels resume', async () => {
  let workbookInfoCall = 0;
  let listModulesCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => {
        workbookInfoCall += 1;
        if (workbookInfoCall === 1) {
          return { success: false, message: 'NO_WORKBOOK: No active workbook.' };
        }
        return {
          success: true,
          name: 'Book2.xlsx',
          path: 'C:\\Book2.xlsx',
          activeSheet: 'Sheet1',
          sheets: ['Sheet1']
        };
      },
      listModules: () => {
        listModulesCalls += 1;
        return { success: true, modules: [{ name: 'Module1' }] };
      }
    }
  });

  const firstResult = await handlers['workbook:info']();
  assert.equal(firstResult.success, false);
  assert.match(firstResult.message, /NO_WORKBOOK/);

  const pausedResult = await handlers['vba:modules']();
  assert.equal(pausedResult.paused, true);
  assert.equal(listModulesCalls, 0);

  const reconnectResult = await handlers['excel:reconnect']();
  assert.equal(reconnectResult.success, true);

  const resumedResult = await handlers['vba:modules']();
  assert.equal(resumedResult.success, true);
  assert.equal(resumedResult.paused, undefined);
  assert.equal(listModulesCalls, 1);
});

test('non-protected channels continue to execute while polling is paused', async () => {
  let selectionCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => ({ success: false, message: 'NO_EXCEL: Excel is not running.' }),
      getSelection: () => {
        selectionCalls += 1;
        return { success: true, address: 'A1', value: 123 };
      }
    }
  });

  const firstResult = await handlers['workbook:info']();
  assert.equal(firstResult.success, false);
  assert.match(firstResult.message, /NO_EXCEL/);

  const selectionResult = await handlers['cell:selection']();
  assert.equal(selectionResult.success, true);
  assert.equal(selectionResult.address, 'A1');
  assert.equal(selectionCalls, 1);
});

test('shortcut audit channels short-circuit while polling is paused', async () => {
  let auditCalls = 0;
  let auditByWorkbookCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => ({ success: false, message: 'NO_EXCEL: Excel is not running.' }),
      auditShortcuts: () => {
        auditCalls += 1;
        return { success: true, shortcuts: [{ macro: 'Module1.RunA', shortcut: 'a' }], unmapped: [] };
      },
      auditShortcutsByWorkbookName: () => {
        auditByWorkbookCalls += 1;
        return { success: true, workbookFound: true, shortcuts: [{ macro: 'Module1.RunA', shortcut: 'a' }], unmapped: [] };
      }
    }
  });

  const firstResult = await handlers['workbook:info']();
  assert.equal(firstResult.success, false);
  assert.match(firstResult.message, /NO_EXCEL/);

  const auditResult = await handlers['vba:shortcut:audit']();
  assert.equal(auditResult.success, false);
  assert.equal(auditResult.paused, true);
  assert.deepEqual(auditResult.shortcuts, []);
  assert.equal(auditCalls, 0);

  const auditByWorkbookResult = await handlers['vba:shortcut:audit:by-workbook'](null, { workbookName: 'Book1.xlsx' });
  assert.equal(auditByWorkbookResult.success, false);
  assert.equal(auditByWorkbookResult.paused, true);
  assert.deepEqual(auditByWorkbookResult.shortcuts, []);
  assert.equal(auditByWorkbookCalls, 0);
});
