const fs = require('node:fs');
const path = require('node:path');

const POLICY_VERSION = 1;
const DEFAULT_POLICY_PATH = path.join(__dirname, 'security-policy.json');

function toTrimmedString(value) {
  return String(value || '').trim();
}

function normalizeStringArray(value) {
  if (!Array.isArray(value)) {
    return [];
  }
  return value
    .map((item) => toTrimmedString(item))
    .filter(Boolean);
}

function toPositiveInt(value, fallback) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < 1) {
    return fallback;
  }
  return Math.floor(parsed);
}

function compileRegexList(patterns) {
  const compiled = [];
  for (const pattern of patterns) {
    compiled.push(new RegExp(pattern, 'i'));
  }
  return compiled;
}

function normalizeRawPolicy(rawPolicy = {}) {
  const version = Number(rawPolicy?.version);
  if (version !== POLICY_VERSION) {
    throw new Error(`Unsupported security policy version "${version}". Expected "${POLICY_VERSION}".`);
  }

  const limits = rawPolicy?.limits || {};
  const allowlists = rawPolicy?.allowlists || {};
  const enforcement = rawPolicy?.enforcement || {};

  const normalized = {
    version: POLICY_VERSION,
    limits: {
      maxVbaCodeChars: toPositiveInt(limits.maxVbaCodeChars, 60000),
      maxModuleNameChars: toPositiveInt(limits.maxModuleNameChars, 80),
      maxMacroNameChars: toPositiveInt(limits.maxMacroNameChars, 255)
    },
    allowlists: {
      modulesExact: normalizeStringArray(allowlists.modulesExact),
      modulesRegex: normalizeStringArray(allowlists.modulesRegex),
      macrosExact: normalizeStringArray(allowlists.macrosExact),
      macrosRegex: normalizeStringArray(allowlists.macrosRegex)
    },
    enforcement: {
      denyByDefault: Boolean(enforcement.denyByDefault)
    }
  };

  const modulesExactSet = new Set(normalized.allowlists.modulesExact.map((entry) => entry.toLowerCase()));
  const macrosExactSet = new Set(normalized.allowlists.macrosExact.map((entry) => entry.toLowerCase()));
  const modulesRegexCompiled = compileRegexList(normalized.allowlists.modulesRegex);
  const macrosRegexCompiled = compileRegexList(normalized.allowlists.macrosRegex);

  return {
    ...normalized,
    compiled: {
      modulesExactSet,
      macrosExactSet,
      modulesRegex: modulesRegexCompiled,
      macrosRegex: macrosRegexCompiled
    }
  };
}

let cachedPolicyPath = '';
let cachedPolicyMtimeMs = 0;
let cachedPolicy = null;

function loadSecurityPolicy(options = {}) {
  const policyPath = path.resolve(toTrimmedString(options?.policyPath) || DEFAULT_POLICY_PATH);
  const forceReload = Boolean(options?.forceReload);

  try {
    const stats = fs.statSync(policyPath);
    const mtimeMs = Number(stats.mtimeMs) || 0;

    if (!forceReload && cachedPolicy && cachedPolicyPath === policyPath && cachedPolicyMtimeMs === mtimeMs) {
      return { ok: true, policy: cachedPolicy, path: policyPath };
    }

    const rawText = fs.readFileSync(policyPath, 'utf8');
    const rawPolicy = JSON.parse(rawText);
    const policy = normalizeRawPolicy(rawPolicy);

    cachedPolicyPath = policyPath;
    cachedPolicyMtimeMs = mtimeMs;
    cachedPolicy = policy;

    return { ok: true, policy, path: policyPath };
  } catch (error) {
    return {
      ok: false,
      reasonCode: 'POLICY_UNAVAILABLE',
      message: 'Security policy is unavailable.',
      error: String(error?.message || error || 'Unknown policy error'),
      path: policyPath
    };
  }
}

function matchesExactOrRegex(value, exactSet, regexList) {
  const normalizedValue = toTrimmedString(value);
  if (!normalizedValue) {
    return false;
  }

  if (exactSet.has(normalizedValue.toLowerCase())) {
    return true;
  }

  return regexList.some((entry) => entry.test(normalizedValue));
}

function isModuleAllowed(policy, moduleName) {
  if (!policy || typeof policy !== 'object') {
    return false;
  }

  return matchesExactOrRegex(
    moduleName,
    policy.compiled?.modulesExactSet || new Set(),
    policy.compiled?.modulesRegex || []
  );
}

function normalizeMacroTargetForPolicy(macroName) {
  const normalized = toTrimmedString(macroName);
  if (!normalized) {
    return '';
  }

  const bangIndex = normalized.indexOf('!');
  if (bangIndex < 0) {
    return normalized;
  }

  return normalized.slice(bangIndex + 1).trim();
}

function isMacroAllowed(policy, macroName) {
  if (!policy || typeof policy !== 'object') {
    return false;
  }

  const normalizedTarget = normalizeMacroTargetForPolicy(macroName);
  return matchesExactOrRegex(
    normalizedTarget,
    policy.compiled?.macrosExactSet || new Set(),
    policy.compiled?.macrosRegex || []
  );
}

module.exports = {
  loadSecurityPolicy,
  isModuleAllowed,
  isMacroAllowed,
  normalizeMacroTargetForPolicy
};
