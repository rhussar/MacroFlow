import assert from 'node:assert/strict';
import test from 'node:test';
import { fetchWorkbookScopedData } from './useWorkbookScopedData.js';

test('fetchWorkbookScopedData normalizes workbook data and can namespace macro identities', async () => {
  const previousWindow = globalThis.window;
  globalThis.window = {
    excel: {
      vba: {
        modulesByWorkbook: async () => ({
          success: true,
          workbookFound: true,
          workbook: { name: 'Target.xlsm', path: 'C:/Target.xlsm' },
          modules: [
            { name: 'Module1', type: 'Standard Module', typeId: 1, lineCount: 12 }
          ]
        }),
        proceduresByWorkbook: async () => ({
          success: true,
          workbookFound: true,
          workbook: { name: 'Target.xlsm', path: 'C:/Target.xlsm' },
          procedures: [
            { module: 'Module1', name: 'RunA', kind: 'Sub', scope: 'Public' }
          ]
        })
      }
    }
  };

  try {
    const result = await fetchWorkbookScopedData(
      { name: 'Target.xlsm', path: 'C:/Target.xlsm' },
      { namespaceMacros: true, defaultWorkbookName: 'Active Workbook' }
    );

    assert.equal(result.success, true);
    assert.equal(result.workbookFound, true);
    assert.equal(result.workbook?.key, 'C:/Target.xlsm');
    assert.deepEqual(result.modules.map((item) => item.id), ['C:/Target.xlsm::module::Module1']);
    assert.deepEqual(result.macros.map((item) => item.id), ['C:/Target.xlsm::macro::Module1::RunA::Sub::Public']);
    assert.deepEqual(result.macros.map((item) => item.fullName), ['Target.xlsm!Module1.RunA']);
  } finally {
    globalThis.window = previousWindow;
  }
});

test('fetchWorkbookScopedData preserves workbook-missing surface without converting it to an error', async () => {
  const previousWindow = globalThis.window;
  globalThis.window = {
    excel: {
      vba: {
        modulesByWorkbook: async () => ({
          success: true,
          workbookFound: false,
          workbook: null,
          modules: []
        }),
        proceduresByWorkbook: async () => ({
          success: true,
          workbookFound: false,
          workbook: null,
          procedures: []
        })
      }
    }
  };

  try {
    const result = await fetchWorkbookScopedData(
      { name: 'Closed.xlsm', path: 'C:/Closed.xlsm' },
      { defaultWorkbookName: 'Workbook' }
    );

    assert.equal(result.success, true);
    assert.equal(result.workbookFound, false);
    assert.equal(result.workbook?.key, 'C:/Closed.xlsm');
    assert.deepEqual(result.modules, []);
    assert.deepEqual(result.macros, []);
  } finally {
    globalThis.window = previousWindow;
  }
});
