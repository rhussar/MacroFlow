function toSafeString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function isUppercaseLetter(letter) {
  return /^[A-Z]$/.test(letter);
}

function isLowercaseLetter(letter) {
  return /^[a-z]$/.test(letter);
}

function isLetter(letter) {
  return isUppercaseLetter(letter) || isLowercaseLetter(letter);
}

export function parseShortcutLetter(rawShortcut) {
  const raw = toSafeString(rawShortcut);
  if (!raw) {
    return '';
  }

  if (isLetter(raw)) {
    return raw;
  }

  const tokens = raw
    .replace(/\s+/g, ' ')
    .split('+')
    .map((token) => token.trim())
    .filter(Boolean);

  if (tokens.length === 0) {
    return '';
  }

  const keyToken = tokens[tokens.length - 1];
  if (!isLetter(keyToken)) {
    return '';
  }

  const modifierTokens = tokens.slice(0, -1).map((token) => token.toLowerCase());
  const unsupportedModifier = modifierTokens.some(
    (token) => token !== 'ctrl' && token !== 'control' && token !== 'shift'
  );
  if (unsupportedModifier) {
    return '';
  }

  const hasShiftModifier = modifierTokens.includes('shift');
  if (hasShiftModifier) {
    return keyToken.toUpperCase();
  }

  return keyToken;
}

export function normalizeShortcutLetterDraft(input) {
  const raw = toSafeString(input);
  if (!raw) {
    return '';
  }

  for (let i = raw.length - 1; i >= 0; i -= 1) {
    const char = raw[i];
    if (isLetter(char)) {
      return char;
    }
  }

  return '';
}

export function formatShortcutPrefix(letter) {
  return isUppercaseLetter(letter) ? 'Ctrl + Shift +' : 'Ctrl +';
}

export function toExcelShortcutKeyFromLetter(letter) {
  const normalized = normalizeShortcutLetterDraft(letter);
  if (!normalized) {
    return '';
  }
  return normalized;
}
