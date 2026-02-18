import assert from 'node:assert/strict';
import test from 'node:test';
import { buildWorkbookShortcutSnapshotKey } from './useWorkbookShortcutState.js';

test('buildWorkbookShortcutSnapshotKey includes workbook identity and stable sorted macro IDs', () => {
  const snapshotKey = buildWorkbookShortcutSnapshotKey(
    { name: 'Model.xlsm', path: 'C:/Model.xlsm' },
    [{ id: 'macro-b' }, { id: 'macro-a' }]
  );

  assert.equal(snapshotKey, 'C:/Model.xlsm::macro-a|macro-b');
});

test('buildWorkbookShortcutSnapshotKey falls back to workbook name', () => {
  const snapshotKey = buildWorkbookShortcutSnapshotKey(
    { name: 'PERSONAL.XLSB', path: '' },
    []
  );

  assert.equal(snapshotKey, 'PERSONAL.XLSB::');
});

