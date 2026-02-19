import assert from 'node:assert/strict';
import test from 'node:test';
import {
  mapSearchError,
  normalizeMacros,
  normalizeModules,
  normalizeWorkbook
} from './search-data.js';

test('mapSearchError maps NO_EXCEL to no_excel state', () => {
  const result = mapSearchError('NO_EXCEL: Excel is not running');
  assert.equal(result.status, 'no_excel');
  assert.equal(result.code, 'NO_EXCEL');
});

test('mapSearchError maps NO_WORKBOOK to no_workbook state', () => {
  const result = mapSearchError('NO_WORKBOOK: No workbook is open');
  assert.equal(result.status, 'no_workbook');
  assert.equal(result.code, 'NO_WORKBOOK');
});

test('mapSearchError maps MULTI_INSTANCE to multi_instance state', () => {
  const result = mapSearchError('MULTI_INSTANCE: Multiple Excel processes detected');
  assert.equal(result.status, 'multi_instance');
  assert.equal(result.code, 'MULTI_INSTANCE');
});

test('mapSearchError maps unknown errors to error and preserves message', () => {
  const result = mapSearchError('Some unexpected failure');
  assert.equal(result.status, 'error');
  assert.equal(result.code, 'UNKNOWN');
  assert.equal(result.message, 'Some unexpected failure');
});

test('normalizeModules excludes runtime and VBA document objects (sheets/workbook)', () => {
  const modules = normalizeModules(
    [
      { name: 'Module1', type: 'Standard Module', typeId: 1, lineCount: 20 },
      { name: 'MacroFlow_Runtime', type: 'Standard Module', typeId: 1, lineCount: 12 },
      { name: 'ClassOne', type: 'Class Module', typeId: 2, lineCount: 5 },
      { name: 'ThisWorkbook', type: 'Document', typeId: 100, lineCount: 10 },
      { name: 'Sheet1', type: 'document', typeId: null, lineCount: 8 }
    ],
    { name: 'Book1.xlsm', path: 'C:/Book1.xlsm' }
  );

  assert.equal(modules.length, 2);
  assert.deepEqual(
    modules.map((item) => item.name),
    ['Module1', 'ClassOne']
  );
});

test('normalizeWorkbook preserves active sheet from workbook info payload', () => {
  const workbook = normalizeWorkbook({
    success: true,
    name: 'Book1.xlsm',
    path: 'C:/Book1.xlsm',
    activeSheet: 'Summary'
  });

  assert.equal(workbook?.name, 'Book1.xlsm');
  assert.equal(workbook?.path, 'C:/Book1.xlsm');
  assert.equal(workbook?.activeSheet, 'Summary');
});

test('normalizeMacros keeps only public or implicit Sub procedures and excludes runtime helper', () => {
  const macros = normalizeMacros([
    { module: 'Module1', name: 'RunA', kind: 'Sub', scope: 'Public' },
    { module: 'Module1', name: 'PrivateSub', kind: 'Sub', scope: 'Private' },
    { module: 'Module1', name: 'FuncA', kind: 'Function', scope: 'Public' },
    { module: 'Module2', name: 'ImplicitSub', kind: 'Sub', scope: 'Implicit' },
    { module: 'MacroFlow_Runtime', name: 'MacroFlow_RunMacro', kind: 'Sub', scope: 'Public' }
  ]);

  assert.equal(macros.length, 2);
  assert.deepEqual(
    macros.map((item) => item.name),
    ['RunA', 'ImplicitSub']
  );
});

test('normalizeMacros produces deterministic id and canonical runTarget/fullName', () => {
  const [macro] = normalizeMacros([
    { module: 'ModuleA', name: 'DoWork', kind: 'Sub', scope: 'Public' }
  ]);

  assert.equal(macro.id, 'ModuleA::DoWork::Sub::Public');
  assert.equal(macro.runTarget, 'ModuleA.DoWork');
  assert.equal(macro.fullName, 'ModuleA.DoWork');
});
