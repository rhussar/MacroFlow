import assert from 'node:assert/strict';
import test from 'node:test';
import {
  canonicalizeMacroIdentity,
  mapAuditConflictsByShortcut,
  mapAuditShortcutsToMacroIds,
  normalizeAuditResponse,
  normalizeShortcutKey,
  toExcelShortcutKey,
  toCompactMacroName
} from './shortcut-audit.js';

test('normalizeShortcutKey formats display text and preserves shift semantics from letter case', () => {
  assert.equal(normalizeShortcutKey('ctrl + shift + a'), 'Ctrl + Shift + A');
  assert.equal(normalizeShortcutKey('CTRL+SHIFT+A'), 'Ctrl + Shift + A');
  assert.equal(normalizeShortcutKey('ctrl+a'), 'Ctrl + a');
  assert.equal(normalizeShortcutKey('a'), 'Ctrl + a');
  assert.equal(normalizeShortcutKey('A'), 'Ctrl + Shift + A');
});

test('toExcelShortcutKey maps display/user input to Excel macro shortcut key casing', () => {
  assert.equal(toExcelShortcutKey('Ctrl + A'), 'A');
  assert.equal(toExcelShortcutKey('Ctrl + a'), 'a');
  assert.equal(toExcelShortcutKey('Ctrl + Shift + A'), 'A');
  assert.equal(toExcelShortcutKey('ctrl+shift+a'), 'A');
  assert.equal(toExcelShortcutKey('a'), 'a');
  assert.equal(toExcelShortcutKey('A'), 'A');
});

test('toCompactMacroName prefers module.procedure from workbook-qualified macro names', () => {
  assert.equal(toCompactMacroName("'Book 1.xlsm'!Module1.RunReport"), 'Module1.RunReport');
  assert.equal(toCompactMacroName('Book2.xlsm!Sheet1.RefreshData'), 'Sheet1.RefreshData');
  assert.equal(toCompactMacroName('Module3.DoWork'), 'Module3.DoWork');
});

test('normalizeAuditResponse maps rows and computes conflicts using normalized shortcut keys', () => {
  const result = normalizeAuditResponse({
    shortcuts: [
      { macro: 'Book1.xlsm!Module1.RunA', shortcut: 'ctrl + shift + a' },
      { macro: 'Book1.xlsm!Module2.RunB', shortcut: 'CTRL+SHIFT+A' },
      { macro: 'Book1.xlsm!Module3.RunC', shortcut: 'Ctrl + Alt + Z' }
    ],
    unmapped: ['Book1.xlsm!Module4.NoShortcut'],
    note: 'Tracked shortcuts only.'
  });

  assert.equal(result.mapped.length, 3);
  assert.deepEqual(
    result.mapped.map((item) => item.shortcutNorm),
    ['Ctrl + Shift + A', 'Ctrl + Shift + A', 'Ctrl + Shift + Alt + Z']
  );
  assert.equal(result.unmapped.length, 1);
  assert.equal(result.unmapped[0].macroCompact, 'Module4.NoShortcut');
  assert.equal(result.conflicts.length, 1);
  assert.equal(result.conflicts[0].shortcutNorm, 'Ctrl + Shift + A');
  assert.equal(result.conflicts[0].entries.length, 2);
  assert.equal(result.note, 'Tracked shortcuts only.');
});

test('normalizeAuditResponse handles empty or non-conflicting data', () => {
  const result = normalizeAuditResponse({
    shortcuts: [{ macro: 'Book1.xlsm!Module1.RunA', shortcut: 'Ctrl+1' }],
    unmapped: []
  });

  assert.equal(result.mapped.length, 1);
  assert.equal(result.conflicts.length, 0);
  assert.equal(result.unmapped.length, 0);
  assert.equal(result.note, '');
});

test('canonicalizeMacroIdentity strips workbook prefix and normalizes case', () => {
  assert.equal(canonicalizeMacroIdentity("'Book One.xlsm'!Module1.RunA"), 'module1.runa');
  assert.equal(canonicalizeMacroIdentity('Book2.xlsm!SHEET1.RefreshData'), 'sheet1.refreshdata');
  assert.equal(canonicalizeMacroIdentity('Module3.DoWork'), 'module3.dowork');
});

test('mapAuditShortcutsToMacroIds maps workbook-qualified audit entries to macro IDs', () => {
  const macros = [
    {
      id: 'Module1::RunA::Sub::Public',
      name: 'RunA',
      module: 'Module1',
      runTarget: 'Module1.RunA',
      fullName: 'Module1.RunA'
    },
    {
      id: 'Module2::RunB::Sub::Public',
      name: 'RunB',
      module: 'Module2',
      runTarget: 'Module2.RunB',
      fullName: 'Module2.RunB'
    }
  ];

  const map = mapAuditShortcutsToMacroIds(
    {
      shortcuts: [
        { macro: 'Book1.xlsm!Module1.RunA', shortcut: 'ctrl + shift + a' },
        { macro: "'Book 1.xlsm'!module2.runb", shortcut: 'CTRL + 2' },
        { macro: 'Book1.xlsm!Module9.Missing', shortcut: 'CTRL+9' }
      ]
    },
    macros
  );

  assert.deepEqual(map, {
    'Module1::RunA::Sub::Public': 'Ctrl + Shift + A',
    'Module2::RunB::Sub::Public': 'Ctrl + 2'
  });
});

test('mapAuditConflictsByShortcut groups macro IDs that share the same letter shortcut', () => {
  const macros = [
    {
      id: 'Module1::RunA::Sub::Public',
      name: 'RunA',
      module: 'Module1',
      runTarget: 'Module1.RunA',
      fullName: 'Module1.RunA'
    },
    {
      id: 'Module2::RunB::Sub::Public',
      name: 'RunB',
      module: 'Module2',
      runTarget: 'Module2.RunB',
      fullName: 'Module2.RunB'
    },
    {
      id: 'Module3::RunC::Sub::Public',
      name: 'RunC',
      module: 'Module3',
      runTarget: 'Module3.RunC',
      fullName: 'Module3.RunC'
    }
  ];

  const conflicts = mapAuditConflictsByShortcut(
    {
      shortcuts: [
        { macro: 'Book1.xlsm!Module1.RunA', shortcut: 'ctrl + shift + a' },
        { macro: "'Book 1.xlsm'!module2.runb", shortcut: 'CTRL+SHIFT+A' },
        { macro: 'Book1.xlsm!Module3.RunC', shortcut: 'Ctrl + c' }
      ]
    },
    macros
  );

  assert.deepEqual(conflicts, {
    A: ['Module1::RunA::Sub::Public', 'Module2::RunB::Sub::Public']
  });
});

test('mapAuditConflictsByShortcut treats ctrl+shift+a and ctrl+a as different shortcuts', () => {
  const macros = [
    {
      id: 'Module1::RunA::Sub::Public',
      name: 'RunA',
      module: 'Module1',
      runTarget: 'Module1.RunA',
      fullName: 'Module1.RunA'
    },
    {
      id: 'Module2::RunB::Sub::Public',
      name: 'RunB',
      module: 'Module2',
      runTarget: 'Module2.RunB',
      fullName: 'Module2.RunB'
    }
  ];

  const conflicts = mapAuditConflictsByShortcut(
    {
      shortcuts: [
        { macro: 'Book1.xlsm!Module1.RunA', shortcut: 'Ctrl + Shift + A' },
        { macro: 'Book1.xlsm!Module2.RunB', shortcut: 'Ctrl + a' }
      ]
    },
    macros
  );

  assert.deepEqual(conflicts, {});
});

test('mapAuditConflictsByShortcut ignores non-letter shortcuts', () => {
  const macros = [
    {
      id: 'Module1::RunA::Sub::Public',
      name: 'RunA',
      module: 'Module1',
      runTarget: 'Module1.RunA',
      fullName: 'Module1.RunA'
    },
    {
      id: 'Module2::RunB::Sub::Public',
      name: 'RunB',
      module: 'Module2',
      runTarget: 'Module2.RunB',
      fullName: 'Module2.RunB'
    }
  ];

  const conflicts = mapAuditConflictsByShortcut(
    {
      shortcuts: [
        { macro: 'Book1.xlsm!Module1.RunA', shortcut: 'Ctrl + 1' },
        { macro: 'Book1.xlsm!Module2.RunB', shortcut: 'Ctrl + 1' }
      ]
    },
    macros
  );

  assert.deepEqual(conflicts, {});
});
