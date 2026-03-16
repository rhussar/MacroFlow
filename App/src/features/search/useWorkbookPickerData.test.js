import assert from 'node:assert/strict';
import test from 'node:test';
import {
  getWorkbookKey,
  normalizeListContextModules,
  sortAllFilesModules,
  sortWorkbooksForPicker,
  resolveSelectedWorkbookKey,
  qualifyMacroFullName,
  namespaceMacrosForWorkbook
} from '../workbooks/workbook-model.js';

test('getWorkbookKey prefers workbook path and falls back to workbook name', () => {
  assert.equal(getWorkbookKey({ name: 'Book1.xlsm', path: 'C:/Book1.xlsm' }), 'C:/Book1.xlsm');
  assert.equal(getWorkbookKey({ name: 'Book2.xlsm', path: '' }), 'Book2.xlsm');
  assert.equal(getWorkbookKey({ name: '', path: '' }), '');
});

test('sortWorkbooksForPicker keeps active workbook first and sorts remaining names', () => {
  const rows = sortWorkbooksForPicker(
    [
      { key: 'c', name: 'Charlie.xlsm', path: 'C:/Charlie.xlsm' },
      { key: 'a', name: 'Alpha.xlsm', path: 'C:/Alpha.xlsm' },
      { key: 'b', name: 'Bravo.xlsm', path: 'C:/Bravo.xlsm' }
    ],
    'b'
  );

  assert.deepEqual(rows.map((row) => row.key), ['b', 'a', 'c']);
});

test('resolveSelectedWorkbookKey keeps selection when still present and falls back safely', () => {
  const workbooks = [
    { key: 'active', name: 'Active.xlsm' },
    { key: 'second', name: 'Second.xlsm' }
  ];

  assert.equal(
    resolveSelectedWorkbookKey({
      requestedKey: 'second',
      workbooks,
      activeWorkbookKey: 'active'
    }),
    'second'
  );

  assert.equal(
    resolveSelectedWorkbookKey({
      requestedKey: 'missing',
      workbooks,
      activeWorkbookKey: 'active'
    }),
    'active'
  );

  assert.equal(
    resolveSelectedWorkbookKey({
      requestedKey: 'missing',
      workbooks: [{ key: 'fallback', name: 'Fallback.xlsm' }],
      activeWorkbookKey: ''
    }),
    'fallback'
  );
});

test('qualifyMacroFullName adds workbook prefix and quotes workbook names with spaces', () => {
  assert.equal(
    qualifyMacroFullName('Model.xlsm', 'Module1.RunReport'),
    'Model.xlsm!Module1.RunReport'
  );
  assert.equal(
    qualifyMacroFullName('My Model.xlsm', 'Module1.RunReport'),
    '\'My Model.xlsm\'!Module1.RunReport'
  );
  assert.equal(
    qualifyMacroFullName("Client'sModel.xlsm", 'Module1.RunReport'),
    '\'Client\'\'sModel.xlsm\'!Module1.RunReport'
  );
});

test('namespaceMacrosForWorkbook namescopes IDs and sets workbook-qualified fullName', () => {
  const rows = namespaceMacrosForWorkbook(
    [
      {
        id: 'Module1::RunA::Sub::Public',
        module: 'Module1',
        name: 'RunA',
        runTarget: 'Module1.RunA',
        fullName: 'Module1.RunA'
      }
    ],
    { name: 'Workbook One.xlsm', path: 'C:/Workbook One.xlsm' }
  );

  assert.equal(rows.length, 1);
  assert.equal(rows[0].id, 'C:/Workbook One.xlsm::macro::Module1::RunA::Sub::Public');
  assert.equal(rows[0].fullName, '\'Workbook One.xlsm\'!Module1.RunA');
});

test('sortAllFilesModules puts active workbook modules first and sorts remaining by workbook then module', () => {
  const rows = sortAllFilesModules(
    [
      { id: 'z', name: 'Zulu', workbookName: 'ClientB.xlsm', workbookPath: 'C:/ClientB.xlsm' },
      { id: 'a2', name: 'AlphaTwo', workbookName: 'Active.xlsm', workbookPath: 'C:/Active.xlsm' },
      { id: 'a1', name: 'AlphaOne', workbookName: 'Active.xlsm', workbookPath: 'C:/Active.xlsm' },
      { id: 'b1', name: 'Bravo', workbookName: 'ClientA.xlsm', workbookPath: 'C:/ClientA.xlsm' }
    ],
    'C:/Active.xlsm'
  );

  assert.deepEqual(rows.map((row) => row.id), ['a1', 'a2', 'b1', 'z']);
});

test('normalizeListContextModules applies workbook namespace for each module row', () => {
  const rows = normalizeListContextModules([
    { name: 'Module1', type: 'Standard Module', typeId: 1, lineCount: 10, workbookName: 'BookA.xlsm', workbookPath: 'C:/BookA.xlsm' },
    { name: 'Module2', type: 'Standard Module', typeId: 1, lineCount: 20, workbookName: 'BookB.xlsm', workbookPath: 'C:/BookB.xlsm' }
  ]);

  assert.equal(rows.length, 2);
  assert.equal(rows[0].id, 'C:/BookA.xlsm::module::Module1');
  assert.equal(rows[1].id, 'C:/BookB.xlsm::module::Module2');
});
