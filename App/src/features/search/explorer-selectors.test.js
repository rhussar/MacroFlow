import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildExplorerTree,
  getDefaultExpandedIds
} from './explorer-selectors.js';

test('buildExplorerTree keeps personal workbook metadata on personal module nodes', () => {
  const tree = buildExplorerTree(
    {
      status: 'ready',
      workbook: { name: 'Active.xlsm', path: 'C:/Active.xlsm' },
      modules: [],
      macros: []
    },
    {
      status: 'ready',
      workbookPath: 'C:/Users/ronan/AppData/Roaming/Microsoft/Excel/XLSTART/PERSONAL.XLSB',
      macros: [
        {
          id: 'personal::macro::Module1::RunA',
          module: 'Module1',
          name: 'RunA',
          workbookName: 'PERSONAL.XLSB',
          workbookPath: 'C:/Users/ronan/AppData/Roaming/Microsoft/Excel/XLSTART/PERSONAL.XLSB'
        }
      ]
    }
  );

  assert.equal(tree.length, 2);
  assert.equal(tree[1].children.length, 1);
  assert.deepEqual(tree[1].children[0].data, {
    name: 'Module1',
    workbookName: 'PERSONAL.XLSB',
    workbookPath: 'C:/Users/ronan/AppData/Roaming/Microsoft/Excel/XLSTART/PERSONAL.XLSB'
  });
});

test('buildExplorerTree includes other open workbooks from list context', () => {
  const tree = buildExplorerTree({
    searchData: {
      status: 'ready',
      workbook: { name: 'Active.xlsm', path: 'C:/Active.xlsm' },
      modules: [
        {
          id: 'active::Module1',
          name: 'Module1',
          workbookName: 'Active.xlsm',
          workbookPath: 'C:/Active.xlsm'
        }
      ],
      macros: []
    },
    personalState: {
      status: 'idle',
      macros: []
    },
    workbooks: [
      { name: 'Book2.xlsm', path: 'C:/Book2.xlsm' },
      { name: 'Book3.xlsm', path: 'C:/Book3.xlsm' }
    ],
    allFilesModules: [
      {
        id: 'book2::ModuleA',
        name: 'ModuleA',
        workbookName: 'Book2.xlsm',
        workbookPath: 'C:/Book2.xlsm'
      },
      {
        id: 'book3::ModuleB',
        name: 'ModuleB',
        workbookName: 'Book3.xlsm',
        workbookPath: 'C:/Book3.xlsm'
      }
    ]
  });

  assert.equal(tree.length, 3);
  assert.deepEqual(
    tree.map((node) => node.label),
    ['Active.xlsm', 'Book2.xlsm', 'Book3.xlsm']
  );
  assert.deepEqual(
    tree[1].children.map((node) => node.label),
    ['ModuleA']
  );
  assert.deepEqual(
    tree[2].children.map((node) => node.label),
    ['ModuleB']
  );
});

test('getDefaultExpandedIds starts all workbook roots collapsed', () => {
  const ids = getDefaultExpandedIds([
    { id: 'wb::active', nodeType: 'workbook', label: 'Active.xlsm', children: [] },
    { id: 'wb::PERSONAL.XLSB', nodeType: 'workbook', label: 'PERSONAL.XLSB (Global Macros)', children: [] }
  ]);

  assert.equal(ids.has('wb::active'), false);
  assert.equal(ids.has('wb::PERSONAL.XLSB'), false);
  assert.equal(ids.size, 0);
});
