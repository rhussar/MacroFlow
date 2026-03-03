const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const {
  loadSecurityPolicy,
  isModuleAllowed,
  isMacroAllowed,
  normalizeMacroTargetForPolicy
} = require('./security-policy');

function createTempPolicyFile(contents) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'macroflow-policy-'));
  const filePath = path.join(dir, 'security-policy.json');
  fs.writeFileSync(filePath, JSON.stringify(contents, null, 2), 'utf8');
  return { dir, filePath };
}

test('loadSecurityPolicy parses a valid policy and compiles allowlists', () => {
  const { dir, filePath } = createTempPolicyFile({
    version: 1,
    limits: {
      maxVbaCodeChars: 60000,
      maxModuleNameChars: 80,
      maxMacroNameChars: 255
    },
    allowlists: {
      modulesExact: ['ClientModule'],
      modulesRegex: ['^MacroFlowModule[0-9]+$'],
      macrosExact: ['ClientModule.RunNow'],
      macrosRegex: ['^MacroFlowModule[0-9]+\\.[A-Za-z_][A-Za-z0-9_]*$']
    },
    enforcement: {
      denyByDefault: true
    }
  });

  try {
    const result = loadSecurityPolicy({ policyPath: filePath, forceReload: true });
    assert.equal(result.ok, true);
    assert.equal(result.policy.version, 1);
    assert.equal(result.policy.limits.maxVbaCodeChars, 60000);
    assert.equal(isModuleAllowed(result.policy, 'clientmodule'), true);
    assert.equal(isModuleAllowed(result.policy, 'MacroFlowModule9'), true);
    assert.equal(isModuleAllowed(result.policy, 'UnapprovedModule'), false);
    assert.equal(isMacroAllowed(result.policy, 'ClientModule.RunNow'), true);
    assert.equal(isMacroAllowed(result.policy, "'Book1.xlsm'!MacroFlowModule9.RunA"), true);
    assert.equal(isMacroAllowed(result.policy, 'UnknownModule.RunA'), false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('loadSecurityPolicy returns POLICY_UNAVAILABLE for invalid policy content', () => {
  const { dir, filePath } = createTempPolicyFile({
    version: 999
  });

  try {
    const result = loadSecurityPolicy({ policyPath: filePath, forceReload: true });
    assert.equal(result.ok, false);
    assert.equal(result.reasonCode, 'POLICY_UNAVAILABLE');
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('normalizeMacroTargetForPolicy strips workbook qualifier', () => {
  assert.equal(
    normalizeMacroTargetForPolicy("'Book One.xlsm'!MacroFlowModule7.RunA"),
    'MacroFlowModule7.RunA'
  );
  assert.equal(
    normalizeMacroTargetForPolicy('MacroFlowModule7.RunA'),
    'MacroFlowModule7.RunA'
  );
});
