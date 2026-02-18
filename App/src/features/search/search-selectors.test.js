import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildMacroRowUiModel,
  getSearchStatusView,
  selectAllFilesModules,
  selectActiveWorkbookMacros,
  selectPersonalGlobalMacros,
  selectPersonalGlobalSectionModel
} from './search-selectors.js';

test('buildMacroRowUiModel creates namespaced UI ids for stable row keys', () => {
  const macro = {
    id: 'Module1::RunA::Sub::Public',
    name: 'RunA',
    module: 'Module1'
  };

  const activeRow = buildMacroRowUiModel(macro, 'active');
  const personalRow = buildMacroRowUiModel(macro, 'personal');

  assert.equal(activeRow.uiId, 'active::Module1::RunA::Sub::Public');
  assert.equal(personalRow.uiId, 'personal::Module1::RunA::Sub::Public');
});

test('selectActiveWorkbookMacros filters by query and shortcut tokens', () => {
  const rows = selectActiveWorkbookMacros(
    [
      { id: 'm1', name: 'ApplyTheme', module: 'StyleModule' },
      { id: 'm2', name: 'CleanData', module: 'DataModule' }
    ],
    'shift a',
    { m1: 'A' }
  );

  assert.equal(rows.length, 1);
  assert.equal(rows[0].source, 'active');
  assert.equal(rows[0].macro.id, 'm1');
});

test('selectPersonalGlobalMacros returns personal rows and query filtering', () => {
  const rows = selectPersonalGlobalMacros(
    [
      { id: 'p1', name: 'EmailPDF', module: 'PersonalUtils' },
      { id: 'p2', name: 'ArchiveRows', module: 'OpsTools' }
    ],
    'personal'
  );

  assert.equal(rows.length, 1);
  assert.equal(rows[0].source, 'personal');
  assert.equal(rows[0].macro.id, 'p1');
});

test('selectPersonalGlobalSectionModel maps workbook-missing to empty-state messaging', () => {
  const section = selectPersonalGlobalSectionModel({
    activeWorkbookName: 'ClientModel.xlsm',
    rows: [],
    status: 'ready',
    workbookFound: false,
    error: null
  });

  assert.equal(section.hidden, false);
  assert.equal(section.count, 0);
  assert.equal(section.isEmpty, true);
  assert.equal(section.emptyMessage, 'PERSONAL.XLSB is not open.');
});

test('selectPersonalGlobalSectionModel hides section when active workbook is PERSONAL.XLSB', () => {
  const section = selectPersonalGlobalSectionModel({
    activeWorkbookName: 'personal.xlsb',
    rows: [{ uiId: 'personal::p1', source: 'personal', macro: { id: 'p1' } }],
    status: 'ready',
    workbookFound: true
  });

  assert.equal(section.hidden, true);
  assert.equal(section.count, 0);
  assert.equal(section.isEmpty, false);
});

test('getSearchStatusView returns loading state metadata', () => {
  const statusView = getSearchStatusView({ status: 'loading' });
  assert.equal(statusView.status, 'loading');
  assert.equal(statusView.title, 'Loading workbook data');
  assert.equal(statusView.isLoading, true);
});

test('getSearchStatusView uses backend error override when present', () => {
  const statusView = getSearchStatusView({
    status: 'error',
    error: { message: 'VBA project access denied.' }
  });
  assert.equal(statusView.status, 'error');
  assert.equal(statusView.message, 'VBA project access denied.');
  assert.equal(statusView.isLoading, false);
});

test('selectAllFilesModules filters modules by module name and workbook name', () => {
  const rows = selectAllFilesModules(
    [
      { id: 'm1', name: 'ModuleA', workbookName: 'ClientA.xlsm' },
      { id: 'm2', name: 'ModuleB', workbookName: 'ClientB.xlsm' }
    ],
    'clienta'
  );

  assert.equal(rows.length, 1);
  assert.equal(rows[0].id, 'm1');
});
