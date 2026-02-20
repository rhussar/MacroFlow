const assert = require('node:assert/strict');
const test = require('node:test');
const path = require('node:path');
const Module = require('node:module');

const IPC_HANDLERS_PATH = path.resolve(__dirname, 'ipc-handlers.js');

function loadHandlers({ excelOverrides = {}, appOverrides = {} } = {}) {
  const originalLoad = Module._load;
  const handlers = {};
  const appEvents = {};
  let clearComCacheCalls = 0;
  let quitCalls = 0;

  const excelStub = {
    clearComCache: () => {
      clearComCacheCalls += 1;
    },
    injectModuleByWorkbookName: () => ({
      success: true,
      workbookFound: true,
      workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx' },
      moduleName: 'MacroFlowModule1',
      message: 'ok'
    }),
    getModuleCodeByWorkbookName: () => ({
      success: true,
      workbookFound: true,
      moduleFound: true,
      workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx' },
      moduleName: 'MacroFlowModule1',
      lineCount: 1,
      hash: 'deadbeef',
      code: 'Option Explicit'
    }),
    getModuleSignatureByWorkbookName: () => ({
      success: true,
      workbookFound: true,
      moduleFound: true,
      workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx' },
      moduleName: 'MacroFlowModule1',
      lineCount: 1,
      hash: 'deadbeef'
    }),
    setModuleCodeByWorkbookName: () => ({
      success: true,
      workbookFound: true,
      moduleFound: true,
      workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx' },
      moduleName: 'MacroFlowModule1',
      lineCount: 1,
      hash: 'deadbeef'
    }),
    renameModuleByWorkbookName: () => ({
      success: true,
      workbookFound: true,
      moduleFound: true,
      renamed: true,
      workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx' },
      previousModuleName: 'Module1',
      moduleName: 'Module2',
      message: 'ok'
    }),
    deleteModuleByWorkbookName: () => ({
      success: true,
      workbookFound: true,
      moduleFound: true,
      deleted: true,
      workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx' },
      moduleName: 'Module1',
      message: 'ok'
    }),
    getWorkbookInfo: () => ({ success: true, name: 'Book1' }),
    listModules: () => ({ success: true, modules: [] }),
    listProcedures: () => ({ success: true, procedures: [] }),
    getOpenWorkbooks: () => ({ success: true, workbooks: [] }),
    getActiveWorkbookContext: () => ({
      success: true,
      workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx', activeSheet: 'Sheet1', sheets: ['Sheet1'] },
      modules: [],
      procedures: [],
      shortcutAudit: { success: true, shortcuts: [], unmapped: [] }
    }),
    getOpenWorkbookListContext: () => ({ success: true, workbooks: [], allFilesModules: [] }),
    getPersonalWorkbookStatus: () => ({
      success: true,
      workbookFound: false,
      workbook: null,
      fileExists: false,
      workbookPath: 'C:\\Users\\Test\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB'
    }),
    openPersonalWorkbook: () => ({
      success: true,
      workbookFound: true,
      opened: true,
      alreadyOpen: false,
      workbook: { name: 'PERSONAL.XLSB', path: 'C:\\Users\\Test\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB' },
      fileExists: true,
      workbookPath: 'C:\\Users\\Test\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB'
    }),
    createPersonalWorkbook: () => ({
      success: true,
      created: true,
      opened: true,
      workbookFound: true,
      workbook: { name: 'PERSONAL.XLSB', path: 'C:\\Users\\Test\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB' },
      fileExists: true,
      workbookPath: 'C:\\Users\\Test\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB'
    }),
    auditShortcuts: () => ({ success: true, shortcuts: [], unmapped: [] }),
    auditShortcutsByWorkbookName: () => ({ success: true, shortcuts: [], unmapped: [], workbookFound: true }),
    _focusHelper: {
      findExcelWithWorkbooks: async () => ({ found: false, reason: 'no-qualifying-instance' })
    },
    getSelection: () => ({ success: true, address: 'A1', value: 'x' }),
    ...excelOverrides
  };

  const electronStub = {
    ipcMain: {
      handle: (channel, handler) => {
        handlers[channel] = handler;
      },
      on: (channel, handler) => {
        handlers[channel] = handler;
      }
    },
    app: {
      quit: () => {
        quitCalls += 1;
      },
      on: (event, handler) => {
        appEvents[event] = handler;
      },
      ...appOverrides
    },
    BrowserWindow: {
      getAllWindows: () => []
    }
  };

  const diagnosticsStub = {
    logger: {
      debug: () => {},
      info: () => {},
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
    getQuitCalls: () => quitCalls,
    getClearComCacheCalls: () => clearComCacheCalls,
    triggerAppEvent: (event) => {
      if (typeof appEvents[event] === 'function') {
        appEvents[event]();
      }
    }
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

test('vba:inject:by-workbook forwards workbook args to bridge', async () => {
  let capturedArgs = null;
  const { handlers } = loadHandlers({
    excelOverrides: {
      injectModuleByWorkbookName: (workbookName, moduleName, code, options = {}) => {
        capturedArgs = { workbookName, moduleName, code, options };
        return {
          success: true,
          workbookFound: true,
          workbook: { name: workbookName, path: options.workbookPath || '' },
          moduleName,
          message: 'ok'
        };
      }
    }
  });

  const result = await handlers['vba:inject:by-workbook'](null, {
    workbookName: 'Book1.xlsm',
    workbookPath: 'C:\\Book1.xlsm',
    moduleName: 'MacroFlowModule4',
    code: 'Sub RunA()\nEnd Sub',
    createIfMissing: true
  });

  assert.equal(result.success, true);
  assert.equal(result.workbookFound, true);
  assert.equal(result.moduleName, 'MacroFlowModule4');
  assert.deepEqual(capturedArgs, {
    workbookName: 'Book1.xlsm',
    moduleName: 'MacroFlowModule4',
    code: 'Sub RunA()\nEnd Sub',
    options: {
      workbookPath: 'C:\\Book1.xlsm',
      createIfMissing: true
    }
  });
});

test('module code channels forward workbook args and payloads to bridge', async () => {
  let capturedReadArgs = null;
  let capturedSignatureArgs = null;
  let capturedWriteArgs = null;

  const { handlers } = loadHandlers({
    excelOverrides: {
      getModuleCodeByWorkbookName: (workbookName, moduleName, options = {}) => {
        capturedReadArgs = { workbookName, moduleName, options };
        return {
          success: true,
          workbookFound: true,
          moduleFound: true,
          workbook: { name: workbookName, path: options.workbookPath || '' },
          moduleName,
          lineCount: 2,
          hash: '11111111',
          code: 'Option Explicit\nSub RunA()\nEnd Sub'
        };
      },
      getModuleSignatureByWorkbookName: (workbookName, moduleName, options = {}) => {
        capturedSignatureArgs = { workbookName, moduleName, options };
        return {
          success: true,
          workbookFound: true,
          moduleFound: true,
          workbook: { name: workbookName, path: options.workbookPath || '' },
          moduleName,
          lineCount: 2,
          hash: '22222222'
        };
      },
      setModuleCodeByWorkbookName: (workbookName, moduleName, code, options = {}) => {
        capturedWriteArgs = { workbookName, moduleName, code, options };
        return {
          success: true,
          workbookFound: true,
          moduleFound: true,
          workbook: { name: workbookName, path: options.workbookPath || '' },
          moduleName,
          lineCount: 3,
          hash: '33333333'
        };
      }
    }
  });

  const readResult = await handlers['vba:module-code:by-workbook'](null, {
    workbookName: 'Client.xlsm',
    workbookPath: 'C:\\Client.xlsm',
    moduleName: 'MacroFlowModule7'
  });
  assert.equal(readResult.success, true);
  assert.equal(readResult.hash, '11111111');

  const signatureResult = await handlers['vba:module-signature:by-workbook'](null, {
    workbookName: 'Client.xlsm',
    workbookPath: 'C:\\Client.xlsm',
    moduleName: 'MacroFlowModule7'
  });
  assert.equal(signatureResult.success, true);
  assert.equal(signatureResult.hash, '22222222');

  const writeResult = await handlers['vba:module-code:set:by-workbook'](null, {
    workbookName: 'Client.xlsm',
    workbookPath: 'C:\\Client.xlsm',
    moduleName: 'MacroFlowModule7',
    code: 'Option Explicit',
    createIfMissing: true
  });
  assert.equal(writeResult.success, true);
  assert.equal(writeResult.hash, '33333333');

  assert.deepEqual(capturedReadArgs, {
    workbookName: 'Client.xlsm',
    moduleName: 'MacroFlowModule7',
    options: { workbookPath: 'C:\\Client.xlsm' }
  });
  assert.deepEqual(capturedSignatureArgs, {
    workbookName: 'Client.xlsm',
    moduleName: 'MacroFlowModule7',
    options: { workbookPath: 'C:\\Client.xlsm' }
  });
  assert.deepEqual(capturedWriteArgs, {
    workbookName: 'Client.xlsm',
    moduleName: 'MacroFlowModule7',
    code: 'Option Explicit',
    options: { workbookPath: 'C:\\Client.xlsm', createIfMissing: true }
  });
});

test('module code channels preserve workbook/module not found surfaces', async () => {
  const { handlers } = loadHandlers({
    excelOverrides: {
      getModuleCodeByWorkbookName: () => ({
        success: true,
        workbookFound: false,
        moduleFound: false,
        workbook: null,
        moduleName: 'MacroFlowModule2',
        lineCount: 0,
        hash: '811c9dc5',
        code: '',
        message: 'Workbook \"Missing.xlsm\" is not open.'
      }),
      getModuleSignatureByWorkbookName: () => ({
        success: true,
        workbookFound: true,
        moduleFound: false,
        workbook: { name: 'Client.xlsm', path: 'C:\\Client.xlsm' },
        moduleName: 'MacroFlowModule2',
        lineCount: 0,
        hash: '811c9dc5',
        message: 'Module \"MacroFlowModule2\" was not found.'
      })
    }
  });

  const readResult = await handlers['vba:module-code:by-workbook'](null, {
    workbookName: 'Missing.xlsm',
    workbookPath: 'C:\\Missing.xlsm',
    moduleName: 'MacroFlowModule2'
  });
  assert.equal(readResult.success, true);
  assert.equal(readResult.workbookFound, false);
  assert.equal(readResult.moduleFound, false);

  const signatureResult = await handlers['vba:module-signature:by-workbook'](null, {
    workbookName: 'Client.xlsm',
    workbookPath: 'C:\\Client.xlsm',
    moduleName: 'MacroFlowModule2'
  });
  assert.equal(signatureResult.success, true);
  assert.equal(signatureResult.workbookFound, true);
  assert.equal(signatureResult.moduleFound, false);
});

test('module rename/delete channels forward workbook args and payloads to bridge', async () => {
  let capturedRenameArgs = null;
  let capturedDeleteArgs = null;

  const { handlers } = loadHandlers({
    excelOverrides: {
      renameModuleByWorkbookName: (workbookName, moduleName, nextModuleName, options = {}) => {
        capturedRenameArgs = { workbookName, moduleName, nextModuleName, options };
        return {
          success: true,
          workbookFound: true,
          moduleFound: true,
          renamed: true,
          workbook: { name: workbookName, path: options.workbookPath || '' },
          previousModuleName: moduleName,
          moduleName: nextModuleName
        };
      },
      deleteModuleByWorkbookName: (workbookName, moduleName, options = {}) => {
        capturedDeleteArgs = { workbookName, moduleName, options };
        return {
          success: true,
          workbookFound: true,
          moduleFound: true,
          deleted: true,
          workbook: { name: workbookName, path: options.workbookPath || '' },
          moduleName
        };
      }
    }
  });

  const renameResult = await handlers['vba:module:rename:by-workbook'](null, {
    workbookName: 'Client.xlsm',
    workbookPath: 'C:\\Client.xlsm',
    moduleName: 'OldModule',
    nextModuleName: 'NewModule'
  });
  assert.equal(renameResult.success, true);
  assert.equal(renameResult.renamed, true);

  const deleteResult = await handlers['vba:module:delete:by-workbook'](null, {
    workbookName: 'Client.xlsm',
    workbookPath: 'C:\\Client.xlsm',
    moduleName: 'DeleteMe'
  });
  assert.equal(deleteResult.success, true);
  assert.equal(deleteResult.deleted, true);

  assert.deepEqual(capturedRenameArgs, {
    workbookName: 'Client.xlsm',
    moduleName: 'OldModule',
    nextModuleName: 'NewModule',
    options: { workbookPath: 'C:\\Client.xlsm' }
  });

  assert.deepEqual(capturedDeleteArgs, {
    workbookName: 'Client.xlsm',
    moduleName: 'DeleteMe',
    options: { workbookPath: 'C:\\Client.xlsm' }
  });
});

test('module code channels are not short-circuited by polling pause', async () => {
  let readCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => ({ success: false, message: 'NO_EXCEL: Excel is not running.' }),
      getModuleCodeByWorkbookName: () => {
        readCalls += 1;
        return {
          success: true,
          workbookFound: true,
          moduleFound: true,
          workbook: { name: 'Client.xlsm', path: 'C:\\Client.xlsm' },
          moduleName: 'MacroFlowModule4',
          lineCount: 1,
          hash: 'abcd1234',
          code: 'Option Explicit'
        };
      }
    }
  });

  const firstResult = await handlers['workbook:info']();
  assert.equal(firstResult.success, false);
  assert.match(firstResult.message, /NO_EXCEL/);

  const readResult = await handlers['vba:module-code:by-workbook'](null, {
    workbookName: 'Client.xlsm',
    workbookPath: 'C:\\Client.xlsm',
    moduleName: 'MacroFlowModule4'
  });
  assert.equal(readResult.success, true);
  assert.equal(readCalls, 1);
});

test('workbook context channels short-circuit while polling is paused', async () => {
  let contextCalls = 0;
  let listContextCalls = 0;

  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => ({ success: false, message: 'NO_EXCEL: Excel is not running.' }),
      getActiveWorkbookContext: () => {
        contextCalls += 1;
        return {
          success: true,
          workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx', activeSheet: 'Sheet1', sheets: ['Sheet1'] },
          modules: [{ name: 'Module1' }],
          procedures: [{ name: 'RunA' }],
          shortcutAudit: { success: true, shortcuts: [], unmapped: [] }
        };
      },
      getOpenWorkbookListContext: () => {
        listContextCalls += 1;
        return {
          success: true,
          workbooks: [{ name: 'Book1.xlsx', path: 'C:\\Book1.xlsx' }],
          allFilesModules: [{ name: 'Module1', workbookName: 'Book1.xlsx', workbookPath: 'C:\\Book1.xlsx' }]
        };
      }
    }
  });

  const firstResult = await handlers['workbook:info']();
  assert.equal(firstResult.success, false);
  assert.match(firstResult.message, /NO_EXCEL/);

  const contextResult = await handlers['workbook:context']();
  assert.equal(contextResult.success, false);
  assert.equal(contextResult.paused, true);
  assert.deepEqual(contextResult.modules, []);
  assert.equal(contextCalls, 0);

  const listContextResult = await handlers['workbook:list-context']();
  assert.equal(listContextResult.success, false);
  assert.equal(listContextResult.paused, true);
  assert.deepEqual(listContextResult.workbooks, []);
  assert.equal(listContextCalls, 0);
});

test('workbook:context dedupes concurrent in-flight calls', async () => {
  let contextCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getActiveWorkbookContext: async () => {
        contextCalls += 1;
        await new Promise((resolve) => setTimeout(resolve, 25));
        return {
          success: true,
          workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx', activeSheet: 'Sheet1', sheets: ['Sheet1'] },
          modules: [{ name: 'Module1' }],
          procedures: [{ name: 'RunA' }],
          shortcutAudit: { success: true, shortcuts: [], unmapped: [] }
        };
      }
    }
  });

  const [first, second] = await Promise.all([
    handlers['workbook:context'](),
    handlers['workbook:context']()
  ]);

  assert.equal(contextCalls, 1);
  assert.equal(first.success, true);
  assert.equal(second.success, true);
  assert.equal(first.workbook?.name, 'Book1.xlsx');
  assert.equal(second.workbook?.name, 'Book1.xlsx');
});

test('workbook:context burst cache reuses result for 1s and refreshes after expiry', async () => {
  let contextCalls = 0;
  const originalNow = Date.now;
  let fakeNow = 10_000;
  Date.now = () => fakeNow;

  try {
    const { handlers } = loadHandlers({
      excelOverrides: {
        getActiveWorkbookContext: () => {
          contextCalls += 1;
          return {
            success: true,
            workbook: { name: `Book${contextCalls}.xlsx`, path: `C:\\Book${contextCalls}.xlsx`, activeSheet: 'Sheet1', sheets: ['Sheet1'] },
            modules: [{ name: `Module${contextCalls}` }],
            procedures: [{ name: `Run${contextCalls}` }],
            shortcutAudit: { success: true, shortcuts: [], unmapped: [] }
          };
        }
      }
    });

    const first = await handlers['workbook:context']();
    assert.equal(first.success, true);
    assert.equal(contextCalls, 1);

    fakeNow += 500;
    const second = await handlers['workbook:context']();
    assert.equal(second.success, true);
    assert.equal(contextCalls, 1);
    assert.equal(second.workbook?.name, first.workbook?.name);

    fakeNow += 1001;
    const third = await handlers['workbook:context']();
    assert.equal(third.success, true);
    assert.equal(contextCalls, 2);
    assert.notEqual(third.workbook?.name, first.workbook?.name);
  } finally {
    Date.now = originalNow;
  }
});

test('workbook:context burst cache invalidates after reconnect and resolve success', async () => {
  let contextCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => ({
        success: true,
        name: 'Book2.xlsx',
        path: 'C:\\Book2.xlsx',
        activeSheet: 'Sheet1',
        sheets: ['Sheet1']
      }),
      getActiveWorkbookContext: () => {
        contextCalls += 1;
        return {
          success: true,
          workbook: { name: `Book${contextCalls}.xlsx`, path: `C:\\Book${contextCalls}.xlsx`, activeSheet: 'Sheet1', sheets: ['Sheet1'] },
          modules: [],
          procedures: [],
          shortcutAudit: { success: true, shortcuts: [], unmapped: [] }
        };
      }
    }
  });

  const first = await handlers['workbook:context']();
  assert.equal(first.success, true);
  assert.equal(contextCalls, 1);

  const cached = await handlers['workbook:context']();
  assert.equal(cached.success, true);
  assert.equal(contextCalls, 1);

  const reconnect = await handlers['excel:reconnect']();
  assert.equal(reconnect.success, true);
  const afterReconnect = await handlers['workbook:context']();
  assert.equal(afterReconnect.success, true);
  assert.equal(contextCalls, 2);

  const resolve = await handlers['excel:resolveInstance']();
  assert.equal(resolve.resolved, true);
  const afterResolve = await handlers['workbook:context']();
  assert.equal(afterResolve.success, true);
  assert.equal(contextCalls, 3);
});

test('personal channels route to bridge and return payloads', async () => {
  let statusCalls = 0;
  let openCalls = 0;
  let createCalls = 0;

  const { handlers } = loadHandlers({
    excelOverrides: {
      getPersonalWorkbookStatus: () => {
        statusCalls += 1;
        return {
          success: true,
          workbookFound: false,
          workbook: null,
          fileExists: false,
          workbookPath: 'C:\\XLSTART\\PERSONAL.XLSB'
        };
      },
      openPersonalWorkbook: () => {
        openCalls += 1;
        return {
          success: true,
          workbookFound: true,
          opened: true,
          alreadyOpen: false,
          workbook: { name: 'PERSONAL.XLSB', path: 'C:\\XLSTART\\PERSONAL.XLSB' },
          fileExists: true,
          workbookPath: 'C:\\XLSTART\\PERSONAL.XLSB'
        };
      },
      createPersonalWorkbook: () => {
        createCalls += 1;
        return {
          success: true,
          created: true,
          opened: true,
          workbookFound: true,
          workbook: { name: 'PERSONAL.XLSB', path: 'C:\\XLSTART\\PERSONAL.XLSB' },
          fileExists: true,
          workbookPath: 'C:\\XLSTART\\PERSONAL.XLSB'
        };
      }
    }
  });

  const statusResult = await handlers['personal:status']();
  assert.equal(statusResult.success, true);
  assert.equal(statusResult.fileExists, false);
  assert.equal(statusCalls, 1);

  const openResult = await handlers['personal:open']();
  assert.equal(openResult.success, true);
  assert.equal(openResult.opened, true);
  assert.equal(openCalls, 1);

  const createResult = await handlers['personal:create']();
  assert.equal(createResult.success, true);
  assert.equal(createResult.created, true);
  assert.equal(createCalls, 1);
});

test('workbook:context burst cache invalidates after personal:open success', async () => {
  let contextCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getActiveWorkbookContext: () => {
        contextCalls += 1;
        return {
          success: true,
          workbook: { name: `Book${contextCalls}.xlsx`, path: `C:\\Book${contextCalls}.xlsx`, activeSheet: 'Sheet1', sheets: ['Sheet1'] },
          modules: [],
          procedures: [],
          shortcutAudit: { success: true, shortcuts: [], unmapped: [] }
        };
      },
      openPersonalWorkbook: () => ({
        success: true,
        workbookFound: true,
        opened: true,
        alreadyOpen: false,
        workbook: { name: 'PERSONAL.XLSB', path: 'C:\\XLSTART\\PERSONAL.XLSB' },
        fileExists: true,
        workbookPath: 'C:\\XLSTART\\PERSONAL.XLSB'
      })
    }
  });

  const first = await handlers['workbook:context']();
  assert.equal(first.success, true);
  assert.equal(contextCalls, 1);

  const cached = await handlers['workbook:context']();
  assert.equal(cached.success, true);
  assert.equal(contextCalls, 1);

  const openResult = await handlers['personal:open']();
  assert.equal(openResult.success, true);

  const afterOpen = await handlers['workbook:context']();
  assert.equal(afterOpen.success, true);
  assert.equal(contextCalls, 2);
});

test('workbook:context burst cache invalidates after personal:create success', async () => {
  let contextCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getActiveWorkbookContext: () => {
        contextCalls += 1;
        return {
          success: true,
          workbook: { name: `Book${contextCalls}.xlsx`, path: `C:\\Book${contextCalls}.xlsx`, activeSheet: 'Sheet1', sheets: ['Sheet1'] },
          modules: [],
          procedures: [],
          shortcutAudit: { success: true, shortcuts: [], unmapped: [] }
        };
      },
      createPersonalWorkbook: () => ({
        success: true,
        created: true,
        opened: true,
        workbookFound: true,
        workbook: { name: 'PERSONAL.XLSB', path: 'C:\\XLSTART\\PERSONAL.XLSB' },
        fileExists: true,
        workbookPath: 'C:\\XLSTART\\PERSONAL.XLSB'
      })
    }
  });

  const first = await handlers['workbook:context']();
  assert.equal(first.success, true);
  assert.equal(contextCalls, 1);

  const cached = await handlers['workbook:context']();
  assert.equal(cached.success, true);
  assert.equal(contextCalls, 1);

  const createResult = await handlers['personal:create']();
  assert.equal(createResult.success, true);

  const afterCreate = await handlers['workbook:context']();
  assert.equal(afterCreate.success, true);
  assert.equal(contextCalls, 2);
});

test('workbook:context burst cache invalidates after module rename/delete success', async () => {
  let contextCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getActiveWorkbookContext: () => {
        contextCalls += 1;
        return {
          success: true,
          workbook: { name: `Book${contextCalls}.xlsx`, path: `C:\\Book${contextCalls}.xlsx`, activeSheet: 'Sheet1', sheets: ['Sheet1'] },
          modules: [],
          procedures: [],
          shortcutAudit: { success: true, shortcuts: [], unmapped: [] }
        };
      },
      renameModuleByWorkbookName: () => ({
        success: true,
        workbookFound: true,
        moduleFound: true,
        renamed: true,
        workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx' },
        previousModuleName: 'OldModule',
        moduleName: 'NewModule'
      }),
      deleteModuleByWorkbookName: () => ({
        success: true,
        workbookFound: true,
        moduleFound: true,
        deleted: true,
        workbook: { name: 'Book1.xlsx', path: 'C:\\Book1.xlsx' },
        moduleName: 'DeleteMe'
      })
    }
  });

  const first = await handlers['workbook:context']();
  assert.equal(first.success, true);
  assert.equal(contextCalls, 1);

  const cached = await handlers['workbook:context']();
  assert.equal(cached.success, true);
  assert.equal(contextCalls, 1);

  const renameResult = await handlers['vba:module:rename:by-workbook'](null, {
    workbookName: 'Book1.xlsx',
    workbookPath: 'C:\\Book1.xlsx',
    moduleName: 'OldModule',
    nextModuleName: 'NewModule'
  });
  assert.equal(renameResult.success, true);
  assert.equal(renameResult.renamed, true);

  const afterRename = await handlers['workbook:context']();
  assert.equal(afterRename.success, true);
  assert.equal(contextCalls, 2);

  const deleteResult = await handlers['vba:module:delete:by-workbook'](null, {
    workbookName: 'Book1.xlsx',
    workbookPath: 'C:\\Book1.xlsx',
    moduleName: 'DeleteMe'
  });
  assert.equal(deleteResult.success, true);
  assert.equal(deleteResult.deleted, true);

  const afterDelete = await handlers['workbook:context']();
  assert.equal(afterDelete.success, true);
  assert.equal(contextCalls, 3);
});

test('excel:resolveInstance dedupes concurrent requests and reuses one helper call', async () => {
  let workbookInfoCalls = 0;
  let helperCalls = 0;

  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => {
        workbookInfoCalls += 1;
        if (workbookInfoCalls === 1) {
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
      _focusHelper: {
        findExcelWithWorkbooks: async () => {
          helperCalls += 1;
          await new Promise((resolve) => setTimeout(resolve, 25));
          return {
            found: true,
            activated: true,
            pid: 10840,
            workbookCount: 1,
            strategy: 'foreground'
          };
        }
      }
    }
  });

  const [first, second] = await Promise.all([
    handlers['excel:resolveInstance'](),
    handlers['excel:resolveInstance']()
  ]);

  assert.equal(helperCalls, 1);
  assert.equal(workbookInfoCalls, 2);
  assert.equal(first.resolved, true);
  assert.equal(second.resolved, true);
  assert.equal(first.reason, 'helper-resolved');
  assert.equal(first.pid, 10840);
  assert.equal(first.workbookCount, 1);
  assert.equal(first.strategy, 'foreground');
  assert.equal(second.pid, 10840);
});

test('excel:resolveInstance returns helper metadata when no qualifying instance is found', async () => {
  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => ({ success: false, message: 'NO_WORKBOOK: No active workbook.' }),
      _focusHelper: {
        findExcelWithWorkbooks: async () => ({
          found: false,
          reason: 'timeout',
          pid: 2222,
          workbookCount: 0,
          strategy: 'max_workbooks'
        })
      }
    }
  });

  const result = await handlers['excel:resolveInstance']();
  assert.equal(result.resolved, false);
  assert.equal(result.reason, 'timeout');
  assert.equal(result.pid, 2222);
  assert.equal(result.workbookCount, 0);
  assert.equal(result.strategy, 'max_workbooks');
});

test('before-quit latch blocks Excel COM calls during shutdown', async () => {
  let workbookInfoCalls = 0;
  const { handlers, triggerAppEvent } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => {
        workbookInfoCalls += 1;
        return { success: true, name: 'Book1.xlsx' };
      }
    }
  });

  triggerAppEvent('before-quit');

  const result = await handlers['workbook:info']();
  assert.equal(result.success, false);
  assert.match(result.message, /APP_SHUTTING_DOWN/);
  assert.equal(workbookInfoCalls, 0);
});

test('app:close sets shutdown latch and eventually calls app.quit', async () => {
  const { handlers, getQuitCalls } = loadHandlers({
    excelOverrides: {
      setShuttingDown: () => {}
    }
  });

  handlers['app:close']();
  await new Promise((resolve) => setTimeout(resolve, 120));
  assert.equal(getQuitCalls() > 0, true);
});
