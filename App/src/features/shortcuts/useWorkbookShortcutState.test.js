import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildWorkbookShortcutSnapshotKey,
  shouldDelayInitialShortcutLoad
} from './useShortcutState.js';

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

test('shouldDelayInitialShortcutLoad only delays the first uncached audit', () => {
  assert.equal(
    shouldDelayInitialShortcutLoad({
      delayMs: 1200,
      hasLoadedShortcuts: false,
      hasSeededAudit: false,
      hasFreshCachedSnapshot: false
    }),
    true
  );

  assert.equal(
    shouldDelayInitialShortcutLoad({
      delayMs: 1200,
      hasLoadedShortcuts: true,
      hasSeededAudit: false,
      hasFreshCachedSnapshot: false
    }),
    false
  );

  assert.equal(
    shouldDelayInitialShortcutLoad({
      delayMs: 1200,
      hasLoadedShortcuts: false,
      hasSeededAudit: true,
      hasFreshCachedSnapshot: false
    }),
    false
  );
});

