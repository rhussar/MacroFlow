import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveInitialModuleNode } from './explorer-selection.js';

test('resolveInitialModuleNode prefers module id when available', () => {
  const tree = [
    {
      id: 'wb::A',
      nodeType: 'workbook',
      label: 'A',
      children: [
        { id: 'A::module::One', nodeType: 'module', label: 'ModuleOne', data: { name: 'ModuleOne' }, children: [] },
        { id: 'A::module::Two', nodeType: 'module', label: 'ModuleTwo', data: { name: 'ModuleTwo' }, children: [] }
      ]
    }
  ];

  const node = resolveInitialModuleNode(tree, {
    moduleId: 'A::module::Two',
    moduleName: 'ModuleOne'
  });

  assert.ok(node);
  assert.equal(node.id, 'A::module::Two');
});

test('resolveInitialModuleNode falls back to module name case-insensitively', () => {
  const tree = [
    {
      id: 'wb::B',
      nodeType: 'workbook',
      label: 'B',
      children: [
        { id: 'B::module::One', nodeType: 'module', label: 'ModuleOne', data: { name: 'ModuleOne' }, children: [] }
      ]
    }
  ];

  const node = resolveInitialModuleNode(tree, {
    moduleId: '',
    moduleName: 'moduleone'
  });

  assert.ok(node);
  assert.equal(node.id, 'B::module::One');
});

test('resolveInitialModuleNode returns null when no module matches context', () => {
  const tree = [
    {
      id: 'wb::C',
      nodeType: 'workbook',
      label: 'C',
      children: [
        { id: 'C::module::One', nodeType: 'module', label: 'ModuleOne', data: { name: 'ModuleOne' }, children: [] }
      ]
    }
  ];

  const node = resolveInitialModuleNode(tree, {
    moduleId: 'C::module::Missing',
    moduleName: 'Unknown'
  });

  assert.equal(node, null);
});
