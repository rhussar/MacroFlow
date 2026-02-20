import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildWorkbookModuleRequest,
  canShowModuleContextActions,
  isStandardModule,
  isValidVbaModuleName,
  normalizeModuleName,
  shouldCommitModuleRename
} from './module-actions.js';

test('normalizeModuleName trims and stringifies values', () => {
  assert.equal(normalizeModuleName('  Module1  '), 'Module1');
  assert.equal(normalizeModuleName(null), '');
});

test('isStandardModule checks typeId equals 1', () => {
  assert.equal(isStandardModule({ typeId: 1 }), true);
  assert.equal(isStandardModule({ typeId: 2 }), false);
  assert.equal(isStandardModule({}), false);
});

test('canShowModuleContextActions allows only standard modules with names', () => {
  assert.equal(canShowModuleContextActions({ typeId: 1, name: 'Module1' }), true);
  assert.equal(canShowModuleContextActions({ typeId: 1, name: '' }), false);
  assert.equal(canShowModuleContextActions({ typeId: 2, name: 'Class1' }), false);
});

test('isValidVbaModuleName enforces VBA-safe naming rules', () => {
  assert.equal(isValidVbaModuleName('Module_1'), true);
  assert.equal(isValidVbaModuleName('1Module'), false);
  assert.equal(isValidVbaModuleName('Module-1'), false);
});

test('shouldCommitModuleRename requires a meaningful name change', () => {
  assert.equal(shouldCommitModuleRename({ currentName: 'Module1', nextName: 'Module1' }), false);
  assert.equal(shouldCommitModuleRename({ currentName: 'Module1', nextName: 'module1' }), false);
  assert.equal(shouldCommitModuleRename({ currentName: 'Module1', nextName: 'Module2' }), true);
});

test('buildWorkbookModuleRequest uses module workbook fields and fallback workbook', () => {
  const direct = buildWorkbookModuleRequest({
    name: 'Module1',
    workbookName: 'Client.xlsm',
    workbookPath: 'C:\\Client.xlsm'
  });
  assert.deepEqual(direct, {
    workbookName: 'Client.xlsm',
    workbookPath: 'C:\\Client.xlsm',
    moduleName: 'Module1'
  });

  const fallback = buildWorkbookModuleRequest(
    { name: 'Module2', workbookName: '', workbookPath: '' },
    { name: 'Fallback.xlsm', path: 'C:\\Fallback.xlsm' }
  );
  assert.deepEqual(fallback, {
    workbookName: 'Fallback.xlsm',
    workbookPath: 'C:\\Fallback.xlsm',
    moduleName: 'Module2'
  });
});
