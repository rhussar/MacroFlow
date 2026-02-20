import { canonicalizeMacroIdentity } from '../../lib/shortcut-audit.js';
import { normalizeShortcutLetterDraft, parseShortcutLetter } from '../../lib/shortcut-keybind.js';

const DEFAULT_MODULE_PREFIX = 'MacroFlowModule';

export const BUILD_MODE_SEED_CODE = 'Option Explicit';

export function normalizeBuildWorkbook(workbook) {
  const name = String(workbook?.name || '').trim();
  const path = String(workbook?.path || '').trim();
  const key = String(workbook?.key || path || name).trim();

  if (!name && !path && !key) {
    return null;
  }

  return {
    name: name || 'Active Workbook',
    path,
    key: key || name || path
  };
}

export function toWorkbookRequest(workbook) {
  return {
    workbookName: String(workbook?.name || '').trim(),
    workbookPath: String(workbook?.path || '').trim()
  };
}

export function resolveBuildLaunchMode(mode) {
  const normalized = String(mode || '').trim().toLowerCase();
  return normalized === 'existing_module' ? 'existing_module' : 'new_module';
}

export function shouldUseStrictWorkbook(launchMode) {
  return resolveBuildLaunchMode(launchMode) === 'existing_module';
}

export function resolveExistingModuleName(modules = [], requestedName = '') {
  const target = String(requestedName || '').trim();
  if (!target) {
    return '';
  }

  const source = Array.isArray(modules) ? modules : [];
  const exactMatch = source.find((moduleItem) => String(moduleItem?.name || '').trim() === target);
  if (exactMatch) {
    return String(exactMatch?.name || '').trim();
  }

  const targetLower = target.toLowerCase();
  const caseInsensitiveMatch = source.find(
    (moduleItem) => String(moduleItem?.name || '').trim().toLowerCase() === targetLower
  );
  return caseInsensitiveMatch ? String(caseInsensitiveMatch?.name || '').trim() : '';
}

export function selectNextModuleName(modules = [], prefix = DEFAULT_MODULE_PREFIX) {
  const safePrefix = String(prefix || DEFAULT_MODULE_PREFIX).trim() || DEFAULT_MODULE_PREFIX;
  const names = new Set(
    (Array.isArray(modules) ? modules : [])
      .map((moduleItem) => String(moduleItem?.name || '').trim())
      .filter(Boolean)
      .map((name) => name.toLowerCase())
  );

  let index = 1;
  while (names.has(`${safePrefix}${index}`.toLowerCase())) {
    index += 1;
  }

  return `${safePrefix}${index}`;
}

export function extractPrimaryMacroName(code) {
  const source = String(code || '');
  const macroMatch = source.match(/^\s*(?:Public\s+|Private\s+)?Sub\s+([A-Za-z_][A-Za-z0-9_]*)/im);
  return macroMatch ? String(macroMatch[1] || '').trim() : '';
}

export function qualifyWorkbookNameForRun(workbookName) {
  const safe = String(workbookName || '').trim();
  if (!safe) {
    return '';
  }
  if (/\s/.test(safe) || safe.includes("'")) {
    return `'${safe.replace(/'/g, "''")}'`;
  }
  return safe;
}

export function buildWorkbookQualifiedRunTarget(workbookName, moduleName, macroName) {
  const module = String(moduleName || '').trim();
  const macro = String(macroName || '').trim();
  if (!module || !macro) {
    return '';
  }

  const runTarget = `${module}.${macro}`;
  const qualifiedWorkbook = qualifyWorkbookNameForRun(workbookName);
  if (!qualifiedWorkbook) {
    return runTarget;
  }
  return `${qualifiedWorkbook}!${runTarget}`;
}

export function buildSessionMacroTarget(moduleName, macroName) {
  const module = String(moduleName || '').trim();
  const macro = String(macroName || '').trim();
  if (!module || !macro) {
    return '';
  }
  return `${module}.${macro}`;
}

function toAuditShortcutRows(auditResult) {
  return Array.isArray(auditResult?.shortcuts) ? auditResult.shortcuts : [];
}

function toRowMacroIdentity(row) {
  const fullName = String(row?.macro || row?.macroName || row?.fullName || '').trim();
  return canonicalizeMacroIdentity(fullName);
}

function toRowShortcutLetter(row) {
  const rawShortcut = String(row?.shortcut || row?.shortcutKey || row?.key || '').trim();
  return parseShortcutLetter(rawShortcut);
}

export function findAssignedShortcutLetterForMacro(auditResult, macroTarget) {
  const targetIdentity = canonicalizeMacroIdentity(macroTarget);
  if (!targetIdentity) {
    return '';
  }

  const rows = toAuditShortcutRows(auditResult);
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    if (toRowMacroIdentity(row) !== targetIdentity) {
      continue;
    }

    const shortcutLetter = toRowShortcutLetter(row);
    if (shortcutLetter) {
      return shortcutLetter;
    }
  }

  return '';
}

export function hasShortcutConflictForMacro(auditResult, macroTarget, shortcutLetter) {
  const targetIdentity = canonicalizeMacroIdentity(macroTarget);
  const normalizedShortcut = normalizeShortcutLetterDraft(shortcutLetter);
  if (!targetIdentity || !normalizedShortcut) {
    return false;
  }

  return toAuditShortcutRows(auditResult).some((row) => {
    const rowIdentity = toRowMacroIdentity(row);
    if (!rowIdentity || rowIdentity === targetIdentity) {
      return false;
    }
    return toRowShortcutLetter(row) === normalizedShortcut;
  });
}

export function resolveBuildWorkbookTarget({
  selectedWorkbook = null,
  selectedWorkbookFound = true,
  activeWorkbook = null
} = {}) {
  const normalizedSelected = normalizeBuildWorkbook(selectedWorkbook);
  const normalizedActive = normalizeBuildWorkbook(activeWorkbook);

  if (normalizedSelected && selectedWorkbookFound) {
    return normalizedSelected;
  }

  if (normalizedActive) {
    return normalizedActive;
  }

  return normalizedSelected;
}

export function shouldSyncOnBuildExit({
  hasSession = true,
  hasPendingChanges = false
} = {}) {
  return Boolean(hasSession && hasPendingChanges);
}

export function resolveBuildExitAction(action) {
  const normalized = String(action || '').trim().toLowerCase();
  if (normalized === 'retry') {
    return 'retry';
  }
  if (normalized === 'exit_without_save') {
    return 'exit_without_save';
  }
  return 'cancel';
}
