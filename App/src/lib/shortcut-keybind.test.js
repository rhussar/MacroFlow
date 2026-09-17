import assert from 'node:assert/strict';
import test from 'node:test';
import {
  formatShortcutPrefix,
  normalizeShortcutLetterDraft,
  parseShortcutLetter,
  toExcelShortcutKeyFromLetter
} from './shortcut-keybind.js';

test('parseShortcutLetter extracts letter shortcuts from raw audit/user strings', () => {
  assert.equal(parseShortcutLetter('Ctrl+Shift+A'), 'A');
  assert.equal(parseShortcutLetter('Ctrl + a'), 'a');
  assert.equal(parseShortcutLetter('A'), 'A');
  assert.equal(parseShortcutLetter('a'), 'a');
});

test('parseShortcutLetter rejects unsupported non-letter or unsupported modifier shortcuts', () => {
  assert.equal(parseShortcutLetter('Ctrl+1'), '');
  assert.equal(parseShortcutLetter('Ctrl+Alt+z'), '');
  assert.equal(parseShortcutLetter(''), '');
});

test('formatShortcutPrefix includes Shift for uppercase letters only', () => {
  assert.equal(formatShortcutPrefix('A'), 'Ctrl + Shift +');
  assert.equal(formatShortcutPrefix('a'), 'Ctrl +');
  assert.equal(formatShortcutPrefix(''), 'Ctrl +');
});

test('normalizeShortcutLetterDraft keeps one valid letter and rejects non-letters', () => {
  assert.equal(normalizeShortcutLetterDraft('a'), 'a');
  assert.equal(normalizeShortcutLetterDraft('AB'), 'B');
  assert.equal(normalizeShortcutLetterDraft('x1'), 'x');
  assert.equal(normalizeShortcutLetterDraft('1'), '');
  assert.equal(normalizeShortcutLetterDraft(''), '');
});

test('toExcelShortcutKeyFromLetter preserves letter case for shift semantics', () => {
  assert.equal(toExcelShortcutKeyFromLetter('A'), 'A');
  assert.equal(toExcelShortcutKeyFromLetter('a'), 'a');
  assert.equal(toExcelShortcutKeyFromLetter('9'), '');
});
