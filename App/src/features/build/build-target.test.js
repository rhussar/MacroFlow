import assert from 'node:assert/strict';
import test from 'node:test';
import {
  BUILD_MODE_SEED_CODE,
  normalizeBuildWorkbook,
  selectNextModuleName,
  extractPrimaryMacroName,
  buildWorkbookQualifiedRunTarget,
  buildSessionMacroTarget,
  findAssignedShortcutLetterForMacro,
  hasShortcutConflictForMacro,
  resolveBuildWorkbookTarget,
  shouldSyncOnBuildExit,
  resolveBuildExitAction
} from './build-target.js';

test('normalizeBuildWorkbook prefers key/path and returns null for empty input', () => {
  assert.equal(normalizeBuildWorkbook(null), null);
  assert.deepEqual(
    normalizeBuildWorkbook({ name: 'Client.xlsm', path: 'C:/Client.xlsm' }),
    {
      name: 'Client.xlsm',
      path: 'C:/Client.xlsm',
      key: 'C:/Client.xlsm'
    }
  );
});

test('selectNextModuleName picks the first available MacroFlowModule index', () => {
  assert.equal(
    selectNextModuleName([
      { name: 'Module1' },
      { name: 'MacroFlowModule1' },
      { name: 'MacroFlowModule3' }
    ]),
    'MacroFlowModule2'
  );
});

test('extractPrimaryMacroName returns the first Sub name', () => {
  const code = `
Option Explicit

Private Sub RunA()
End Sub

Public Sub RunB()
End Sub
`;

  assert.equal(extractPrimaryMacroName(code), 'RunA');
  assert.equal(extractPrimaryMacroName('Function TestA()\nEnd Function'), '');
});

test('buildWorkbookQualifiedRunTarget applies workbook quoting when needed', () => {
  assert.equal(
    buildWorkbookQualifiedRunTarget('Model.xlsm', 'MacroFlowModule1', 'RunA'),
    'Model.xlsm!MacroFlowModule1.RunA'
  );
  assert.equal(
    buildWorkbookQualifiedRunTarget('My Model.xlsm', 'MacroFlowModule1', 'RunA'),
    '\'My Model.xlsm\'!MacroFlowModule1.RunA'
  );
});

test('buildSessionMacroTarget joins module and macro names', () => {
  assert.equal(buildSessionMacroTarget('MacroFlowModule9', 'RunA'), 'MacroFlowModule9.RunA');
  assert.equal(buildSessionMacroTarget('', 'RunA'), '');
  assert.equal(buildSessionMacroTarget('MacroFlowModule9', ''), '');
});

test('findAssignedShortcutLetterForMacro matches workbook-qualified and unqualified macro identities', () => {
  const auditResult = {
    shortcuts: [
      { macro: "'Book 1.xlsm'!MacroFlowModule9.RunA", shortcut: 'Ctrl + Shift + A' },
      { macro: 'Book 1.xlsm!MacroFlowModule9.RunB', shortcut: 'Ctrl + b' }
    ]
  };

  assert.equal(
    findAssignedShortcutLetterForMacro(auditResult, 'MacroFlowModule9.RunA'),
    'A'
  );
  assert.equal(
    findAssignedShortcutLetterForMacro(auditResult, "'Book 1.xlsm'!MacroFlowModule9.RunB"),
    'b'
  );
  assert.equal(
    findAssignedShortcutLetterForMacro(auditResult, 'MacroFlowModule9.RunC'),
    ''
  );
});

test('hasShortcutConflictForMacro ignores current macro and detects conflicts for other macros', () => {
  const auditResult = {
    shortcuts: [
      { macro: 'Book1.xlsm!MacroFlowModule9.RunA', shortcut: 'Ctrl + Shift + A' },
      { macro: 'Book1.xlsm!MacroFlowModule9.RunB', shortcut: 'Ctrl + Shift + A' },
      { macro: 'Book1.xlsm!MacroFlowModule9.RunC', shortcut: 'Ctrl + c' }
    ]
  };

  assert.equal(
    hasShortcutConflictForMacro(auditResult, 'MacroFlowModule9.RunA', 'A'),
    true
  );
  assert.equal(
    hasShortcutConflictForMacro(auditResult, 'MacroFlowModule9.RunA', 'a'),
    false
  );
  assert.equal(
    hasShortcutConflictForMacro(auditResult, 'MacroFlowModule9.RunC', 'c'),
    false
  );
});

test('build seed code defaults to Option Explicit', () => {
  assert.equal(BUILD_MODE_SEED_CODE, 'Option Explicit');
});

test('resolveBuildWorkbookTarget prefers selected workbook and falls back to active', () => {
  const selectedWorkbook = {
    name: 'Selected.xlsm',
    path: 'C:/Selected.xlsm'
  };
  const activeWorkbook = {
    name: 'Active.xlsm',
    path: 'C:/Active.xlsm'
  };

  assert.deepEqual(
    resolveBuildWorkbookTarget({
      selectedWorkbook,
      selectedWorkbookFound: true,
      activeWorkbook
    }),
    {
      name: 'Selected.xlsm',
      path: 'C:/Selected.xlsm',
      key: 'C:/Selected.xlsm'
    }
  );

  assert.deepEqual(
    resolveBuildWorkbookTarget({
      selectedWorkbook,
      selectedWorkbookFound: false,
      activeWorkbook
    }),
    {
      name: 'Active.xlsm',
      path: 'C:/Active.xlsm',
      key: 'C:/Active.xlsm'
    }
  );
});

test('shouldSyncOnBuildExit syncs only when session has pending changes', () => {
  assert.equal(
    shouldSyncOnBuildExit({
      hasSession: true,
      hasPendingChanges: true
    }),
    true
  );

  assert.equal(
    shouldSyncOnBuildExit({
      hasSession: true,
      hasPendingChanges: false
    }),
    false
  );

  assert.equal(
    shouldSyncOnBuildExit({
      hasSession: false,
      hasPendingChanges: true
    }),
    false
  );
});

test('resolveBuildExitAction maps retry/exit_without_save/cancel correctly', () => {
  assert.equal(resolveBuildExitAction('retry'), 'retry');
  assert.equal(resolveBuildExitAction('exit_without_save'), 'exit_without_save');
  assert.equal(resolveBuildExitAction('cancel'), 'cancel');
  assert.equal(resolveBuildExitAction('anything-else'), 'cancel');
});
