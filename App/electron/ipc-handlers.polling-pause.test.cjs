const assert = require('node:assert/strict');
const test = require('node:test');
const path = require('node:path');
const Module = require('node:module');

const IPC_HANDLERS_PATH = path.resolve(__dirname, 'ipc-handlers.js');

function loadHandlers({
  excelOverrides = {},
  appOverrides = {},
  openAiOverrides = {},
  localAiOverrides = {},
  securityPolicyOverrides = {},
  auditOverrides = {}
} = {}) {
  const originalLoad = Module._load;
  const handlers = {};
  const appEvents = {};
  let clearComCacheCalls = 0;
  let quitCalls = 0;
  const shellCalls = [];

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
    setPersonalWorkbookVisibility: () => ({
      success: true,
      workbookFound: true,
      workbook: { name: 'PERSONAL.XLSB', path: 'C:\\Users\\Test\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB' },
      fileExists: true,
      workbookPath: 'C:\\Users\\Test\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB',
      windowVisible: false,
      windowHidden: true,
      visibilityChanged: true
    }),
    getPersonalWorkbookLocation: () => ({
      success: true,
      workbookPath: 'C:\\Users\\Test\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB',
      folderPath: 'C:\\Users\\Test\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART',
      fileExists: true
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
    shell: {
      showItemInFolder: (target) => {
        shellCalls.push({ fn: 'showItemInFolder', target });
      },
      openPath: async (target) => {
        shellCalls.push({ fn: 'openPath', target });
        return '';
      }
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

  const openAiStub = {
    generateVba: async () => ({
      success: true,
      code: 'Option Explicit\nSub RunA()\nEnd Sub',
      model: 'claude-sonnet-5'
    }),
    generateVbaStream: async (args, deps, onToken) => {
      const result = await (openAiOverrides.generateVba || openAiStub.generateVba)(args);
      return result;
    },
    ...openAiOverrides
  };

  const localAiStub = {
    getStatus: async () => ({
      success: true,
      provider: 'ollama',
      model: 'qwen2.5-coder:3b',
      ready: true,
      needsSetup: false,
      setupInProgress: false,
      runtimeInstalled: true,
      runtimeCommand: 'C:\\Users\\Test\\AppData\\Local\\Programs\\Ollama\\ollama.exe',
      serverReachable: true,
      modelInstalled: true,
      stage: 'ready',
      statusText: 'Local AI is ready.',
      progress: null
    }),
    setup: async () => ({
      success: true,
      started: true,
      status: {
        success: true,
        provider: 'ollama',
        model: 'qwen2.5-coder:3b',
        ready: false,
        needsSetup: true,
        setupInProgress: true,
        runtimeInstalled: false,
        serverReachable: false,
        modelInstalled: false,
        stage: 'checking',
        statusText: 'Checking local AI runtime...'
      }
    }),
    subscribe: () => () => {},
    ...localAiOverrides
  };

  const securityPolicyStub = {
    loadSecurityPolicy: () => ({
      ok: true,
      policy: {
        limits: {
          maxVbaCodeChars: 60000,
          maxModuleNameChars: 80,
          maxMacroNameChars: 255
        },
        allowlists: {
          modulesExact: [],
          modulesRegex: ['^MacroFlowModule[0-9]+$'],
          macrosExact: [],
          macrosRegex: ['^MacroFlowModule[0-9]+\\.[A-Za-z_][A-Za-z0-9_]*$']
        },
        enforcement: { denyByDefault: true }
      }
    }),
    isModuleAllowed: () => true,
    isMacroAllowed: () => true,
    normalizeMacroTargetForPolicy: (macroName) => String(macroName || '').split('!').pop() || '',
    ...securityPolicyOverrides
  };

  const auditStub = {
    initializeAuditLog: () => ({ success: true, auditDirectoryPath: 'C:\\Audit' }),
    writeAuditEvent: async () => ({ success: true, filePath: 'C:\\Audit\\audit-2026-03-02.jsonl' }),
    ...auditOverrides
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
    if (request === './llm-client' || request === './openai-client') {
      return openAiStub;
    }
    if (request === './local-ai-manager') {
      return localAiStub;
    }
    if (request === './security-policy') {
      return securityPolicyStub;
    }
    if (request === './audit-log') {
      return auditStub;
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
    openAiStub,
    localAiStub,
    securityPolicyStub,
    auditStub,
    getQuitCalls: () => quitCalls,
    getClearComCacheCalls: () => clearComCacheCalls,
    getShellCalls: () => shellCalls.slice(),
    triggerAppEvent: (event) => {
      if (typeof appEvents[event] === 'function') {
        appEvents[event]();
      }
    }
  };
}

async function setSelectedWorkbookScope(handlers, workbookName = 'Book1.xlsx', workbookPath = 'C:\\Book1.xlsx') {
  const result = await handlers['security:set-selected-workbook'](null, { workbookName, workbookPath });
  assert.equal(result.success, true);
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

test('polling pause latches after NO_VISIBLE_WINDOWS and keeps reason alignment', async () => {
  let getWorkbookInfoCalls = 0;
  let listModulesCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => {
        getWorkbookInfoCalls += 1;
        return {
          success: false,
          message: 'NO_VISIBLE_WINDOWS: Excel process found but has no visible workbook windows.'
        };
      },
      listModules: () => {
        listModulesCalls += 1;
        return { success: true, modules: [{ name: 'Module1' }] };
      }
    }
  });

  const workbookInfoResult = await handlers['workbook:info']();
  assert.equal(workbookInfoResult.success, false);
  assert.match(workbookInfoResult.message, /NO_VISIBLE_WINDOWS/);
  assert.equal(getWorkbookInfoCalls, 1);

  const modulesResult = await handlers['vba:modules']();
  assert.equal(modulesResult.success, false);
  assert.equal(modulesResult.paused, true);
  assert.equal(modulesResult.reason, 'polling_paused');
  assert.equal(modulesResult.reasonCode, 'NO_VISIBLE_WINDOWS');
  assert.match(modulesResult.message, /NO_VISIBLE_WINDOWS/);
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

test('excel:reconnect success clears NO_VISIBLE_WINDOWS pause and resumes protected channels', async () => {
  let workbookInfoCall = 0;
  let listModulesCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => {
        workbookInfoCall += 1;
        if (workbookInfoCall === 1) {
          return {
            success: false,
            message: 'NO_VISIBLE_WINDOWS: Excel process found but has no visible workbook windows.'
          };
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
  assert.match(firstResult.message, /NO_VISIBLE_WINDOWS/);

  const pausedResult = await handlers['vba:modules']();
  assert.equal(pausedResult.paused, true);
  assert.equal(pausedResult.reasonCode, 'NO_VISIBLE_WINDOWS');
  assert.equal(listModulesCalls, 0);

  const reconnectResult = await handlers['excel:reconnect']();
  assert.equal(reconnectResult.success, true);

  const resumedResult = await handlers['vba:modules']();
  assert.equal(resumedResult.success, true);
  assert.equal(resumedResult.paused, undefined);
  assert.equal(listModulesCalls, 1);
});

test('excel:reconnect applies its own cooldown after NO_EXCEL failure result', async () => {
  let workbookInfoCalls = 0;
  const { handlers } = loadHandlers({
    excelOverrides: {
      getWorkbookInfo: () => {
        workbookInfoCalls += 1;
        return { success: false, message: 'NO_EXCEL: Excel is not running.' };
      }
    }
  });

  const firstReconnect = await handlers['excel:reconnect']();
  assert.equal(firstReconnect.success, false);
  assert.match(firstReconnect.message, /NO_EXCEL/);
  assert.equal(workbookInfoCalls, 1);

  const secondReconnect = await handlers['excel:reconnect']();
  assert.equal(secondReconnect.success, false);
  assert.match(secondReconnect.message, /NO_EXCEL: Waiting for Excel to restart\./);
  assert.equal(
    workbookInfoCalls,
    1,
    'second reconnect should short-circuit during cooldown instead of re-attaching immediately'
  );
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
  await setSelectedWorkbookScope(handlers, 'Book1.xlsm', 'C:\\Book1.xlsm');

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
  await setSelectedWorkbookScope(handlers, 'Client.xlsm', 'C:\\Client.xlsm');

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
  let contextCalls = 0;
  let contextArgs = null;
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
      getPersonalWorkbookContext: (args) => {
        contextCalls += 1;
        contextArgs = args;
        return {
          success: true,
          workbookFound: true,
          workbook: { name: 'PERSONAL.XLSB', path: 'C:\\XLSTART\\PERSONAL.XLSB' },
          fileExists: true,
          workbookPath: 'C:\\XLSTART\\PERSONAL.XLSB',
          procedures: [{ module: 'GlobalMacros', name: 'RunPersonal', kind: 'Sub', scope: 'Public' }],
          shortcutAudit: { success: true, shortcuts: [], unmapped: [] }
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

  const contextResult = await handlers['personal:context']();
  assert.equal(contextResult.success, true);
  assert.equal(contextResult.workbookFound, true);
  assert.equal(contextResult.procedures.length, 1);
  assert.equal(contextCalls, 1);
  assert.deepEqual(contextArgs, { includeShortcutAudit: true });

  const noAuditContextResult = await handlers['personal:context'](null, { includeShortcutAudit: false });
  assert.equal(noAuditContextResult.success, true);
  assert.equal(contextCalls, 2);
  assert.deepEqual(contextArgs, { includeShortcutAudit: false });

  const openResult = await handlers['personal:open']();
  assert.equal(openResult.success, true);
  assert.equal(openResult.opened, true);
  assert.equal(openCalls, 1);

  const createResult = await handlers['personal:create']();
  assert.equal(createResult.success, true);
  assert.equal(createResult.created, true);
  assert.equal(createCalls, 1);
});

test('personal visibility channels forward visibility args and validate payloads', async () => {
  let openArgs = null;
  let createArgs = null;
  let visibilityArgs = null;

  const { handlers } = loadHandlers({
    excelOverrides: {
      openPersonalWorkbook: (args) => {
        openArgs = args;
        return {
          success: true,
          workbookFound: true,
          opened: true,
          alreadyOpen: false,
          workbook: { name: 'PERSONAL.XLSB', path: 'C:\\XLSTART\\PERSONAL.XLSB' },
          fileExists: true,
          workbookPath: 'C:\\XLSTART\\PERSONAL.XLSB',
          windowVisible: true,
          windowHidden: false
        };
      },
      createPersonalWorkbook: (args) => {
        createArgs = args;
        return {
          success: true,
          created: true,
          opened: true,
          workbookFound: true,
          workbook: { name: 'PERSONAL.XLSB', path: 'C:\\XLSTART\\PERSONAL.XLSB' },
          fileExists: true,
          workbookPath: 'C:\\XLSTART\\PERSONAL.XLSB',
          windowVisible: false,
          windowHidden: true
        };
      },
      setPersonalWorkbookVisibility: (args) => {
        visibilityArgs = args;
        return {
          success: true,
          workbookFound: true,
          workbook: { name: 'PERSONAL.XLSB', path: 'C:\\XLSTART\\PERSONAL.XLSB' },
          fileExists: true,
          workbookPath: 'C:\\XLSTART\\PERSONAL.XLSB',
          windowVisible: false,
          windowHidden: true,
          visibilityChanged: true
        };
      }
    }
  });

  const openResult = await handlers['personal:open'](null, { visible: true });
  assert.equal(openResult.success, true);
  assert.deepEqual(openArgs, { visible: true });

  const createResult = await handlers['personal:create'](null, { visible: false });
  assert.equal(createResult.success, true);
  assert.deepEqual(createArgs, { visible: false });

  const visibilityResult = await handlers['personal:visibility:set'](null, { visible: false });
  assert.equal(visibilityResult.success, true);
  assert.deepEqual(visibilityArgs, { visible: false });

  const invalidResult = await handlers['personal:visibility:set'](null, { visible: 'nope' });
  assert.equal(invalidResult.success, false);
  assert.equal(invalidResult.reasonCode, 'VALIDATION_FAILED');
});

test('personal open-folder routes to shell using workbook path when PERSONAL.XLSB exists', async () => {
  const { handlers, getShellCalls } = loadHandlers();

  const result = await handlers['personal:open-folder']();
  assert.equal(result.success, true);
  assert.equal(result.fileExists, true);

  const shellCalls = getShellCalls();
  assert.deepEqual(shellCalls, [
    {
      fn: 'showItemInFolder',
      target: 'C:\\Users\\Test\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB'
    }
  ]);
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

test('ai:generate-vba forwards payload to the local AI client and returns success shape', async () => {
  const calls = [];
  const { handlers } = loadHandlers({
    openAiOverrides: {
      generateVba: async (args) => {
        calls.push(args);
        return {
          success: true,
          code: 'Option Explicit\nPublic Sub RunA()\nEnd Sub',
          model: 'claude-sonnet-5',
          usage: { promptTokens: 10, completionTokens: 12, totalTokens: 22 }
        };
      }
    }
  });

  const result = await handlers['ai:generate-vba'](null, {
    prompt: 'Generate a macro that formats dates',
    workbookName: 'Book1.xlsm',
    moduleName: 'Module1',
    currentCode: 'Option Explicit'
  });

  assert.equal(calls.length, 1);
  assert.equal(calls[0].prompt, 'Generate a macro that formats dates');
  assert.equal(calls[0].workbookName, 'Book1.xlsm');
  assert.equal(calls[0].moduleName, 'Module1');
  assert.equal(result.success, true);
  assert.match(result.code, /Sub RunA/i);
  assert.equal(result.model, 'claude-sonnet-5');
  assert.equal(result.usage.totalTokens, 22);
});

test('ai:generate-vba propagates failure reason and message', async () => {
  const { handlers } = loadHandlers({
    openAiOverrides: {
      generateVba: async () => ({
        success: false,
        reason: 'AI_NOT_READY',
        message: 'Could not reach the AI service. Check your internet connection and try again.'
      })
    }
  });

  const result = await handlers['ai:generate-vba'](null, {
    prompt: 'x',
    workbookName: 'Book1.xlsm',
    moduleName: 'Module1',
    currentCode: ''
  });

  assert.equal(result.success, false);
  assert.equal(result.reason, 'AI_NOT_READY');
  assert.match(result.message, /could not reach the ai service/i);
});

test('ai:status returns cloud AI readiness snapshot', async () => {
  const { handlers } = loadHandlers();

  const result = await handlers['ai:status']();
  assert.equal(result.success, true);
  assert.equal(result.provider, 'anthropic');
  assert.equal(result.ready, true);
  assert.equal(result.model, 'claude-sonnet-5');
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

test('security:set-selected-workbook stores and clears workbook scope', async () => {
  const { handlers } = loadHandlers();

  const selected = await handlers['security:set-selected-workbook'](null, {
    workbookName: 'Client.xlsm',
    workbookPath: 'C:\\Client.xlsm'
  });
  assert.equal(selected.success, true);
  assert.equal(selected.selected, true);

  const cleared = await handlers['security:set-selected-workbook'](null, {});
  assert.equal(cleared.success, true);
  assert.equal(cleared.selected, false);
});

test('vba:inject is blocked with VALIDATION_FAILED and audited', async () => {
  const auditCalls = [];
  const { handlers } = loadHandlers({
    auditOverrides: {
      writeAuditEvent: async (entry) => {
        auditCalls.push(entry);
        return { success: true, filePath: 'C:\\Audit\\audit-2026-03-02.jsonl' };
      }
    }
  });

  const result = await handlers['vba:inject'](null, {
    moduleName: 'MacroFlowModule1',
    code: 'Option Explicit'
  });

  assert.equal(result.success, false);
  assert.equal(result.reasonCode, 'VALIDATION_FAILED');
  assert.equal(auditCalls.length, 1);
  assert.equal(auditCalls[0].outcome, 'blocked');
});

test('high-risk handlers block when selected workbook scope mismatches', async () => {
  const { handlers } = loadHandlers();
  await setSelectedWorkbookScope(handlers, 'Allowed.xlsm', 'C:\\Allowed.xlsm');

  const result = await handlers['vba:inject:by-workbook'](null, {
    workbookName: 'Other.xlsm',
    workbookPath: 'C:\\Other.xlsm',
    moduleName: 'MacroFlowModule1',
    code: 'Option Explicit'
  });

  assert.equal(result.success, false);
  assert.equal(result.reasonCode, 'WORKBOOK_SCOPE_MISMATCH');
});

test('high-risk handlers block when policy denies module/macro', async () => {
  const { handlers } = loadHandlers({
    securityPolicyOverrides: {
      isModuleAllowed: () => false,
      isMacroAllowed: () => false
    }
  });
  await setSelectedWorkbookScope(handlers, 'Book1.xlsx', 'C:\\Book1.xlsx');

  const injectResult = await handlers['vba:inject:by-workbook'](null, {
    workbookName: 'Book1.xlsx',
    workbookPath: 'C:\\Book1.xlsx',
    moduleName: 'MacroFlowModule1',
    code: 'Option Explicit'
  });
  assert.equal(injectResult.success, false);
  assert.equal(injectResult.reasonCode, 'POLICY_DENIED');

  const runResult = await handlers['vba:run'](null, {
    macroName: "'Book1.xlsx'!MacroFlowModule1.RunA"
  });
  assert.equal(runResult.success, false);
  assert.equal(runResult.reasonCode, 'POLICY_DENIED');
});

test('high-risk handlers block oversized VBA payloads', async () => {
  const { handlers } = loadHandlers({
    securityPolicyOverrides: {
      loadSecurityPolicy: () => ({
        ok: true,
        policy: {
          limits: {
            maxVbaCodeChars: 5,
            maxModuleNameChars: 80,
            maxMacroNameChars: 255
          },
          enforcement: { denyByDefault: true }
        }
      })
    }
  });
  await setSelectedWorkbookScope(handlers, 'Book1.xlsx', 'C:\\Book1.xlsx');

  const result = await handlers['vba:module-code:set:by-workbook'](null, {
    workbookName: 'Book1.xlsx',
    workbookPath: 'C:\\Book1.xlsx',
    moduleName: 'MacroFlowModule1',
    code: 'Option Explicit',
    createIfMissing: true
  });

  assert.equal(result.success, false);
  assert.equal(result.reasonCode, 'VALIDATION_FAILED');
});

test('high-risk success paths write audit records', async () => {
  const auditCalls = [];
  const { handlers } = loadHandlers({
    auditOverrides: {
      writeAuditEvent: async (entry) => {
        auditCalls.push(entry);
        return { success: true, filePath: 'C:\\Audit\\audit-2026-03-02.jsonl' };
      }
    }
  });
  await setSelectedWorkbookScope(handlers, 'Book1.xlsx', 'C:\\Book1.xlsx');

  const result = await handlers['vba:inject:by-workbook'](null, {
    workbookName: 'Book1.xlsx',
    workbookPath: 'C:\\Book1.xlsx',
    moduleName: 'MacroFlowModule1',
    code: 'Option Explicit'
  });

  assert.equal(result.success, true);
  assert.equal(auditCalls.length, 1);
  assert.equal(auditCalls[0].outcome, 'succeeded');
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
