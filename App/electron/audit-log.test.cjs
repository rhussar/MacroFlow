const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const { writeAuditEvent } = require('./audit-log');

test('writeAuditEvent writes JSONL record with required fields and code hash only', async () => {
  const auditRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'macroflow-audit-'));

  try {
    const result = await writeAuditEvent(
      {
        action: 'vba.inject',
        channel: 'vba:inject:by-workbook',
        outcome: 'succeeded',
        reasonCode: 'OK',
        workbook: {
          name: 'Client.xlsm',
          path: 'C:\\Client.xlsm'
        },
        target: {
          moduleName: 'MacroFlowModule1',
          macroName: ''
        },
        payload: {
          code: 'Option Explicit\nSub RunA()\nEnd Sub'
        },
        result: {
          success: true,
          message: 'Module injected.'
        }
      },
      { auditRoot }
    );

    assert.equal(result.success, true);
    assert.ok(result.filePath.endsWith('.jsonl'));

    const lines = fs.readFileSync(result.filePath, 'utf8')
      .split(/\r?\n/)
      .filter(Boolean);
    assert.equal(lines.length, 1);
    const row = JSON.parse(lines[0]);

    assert.equal(typeof row.eventId, 'string');
    assert.equal(row.action, 'vba.inject');
    assert.equal(row.channel, 'vba:inject:by-workbook');
    assert.equal(row.outcome, 'succeeded');
    assert.equal(row.actor.username.length > 0, true);
    assert.equal(row.actor.deviceId.length > 0, true);
    assert.equal(row.workbook.name, 'Client.xlsm');
    assert.equal(row.workbook.path, 'C:\\Client.xlsm');
    assert.equal(typeof row.workbook.identifierHash, 'string');
    assert.equal(row.target.moduleName, 'MacroFlowModule1');
    assert.equal(row.payload.codeLength > 0, true);
    assert.equal(typeof row.payload.codeHash, 'string');
    assert.equal(Object.prototype.hasOwnProperty.call(row.payload, 'code'), false);
  } finally {
    fs.rmSync(auditRoot, { recursive: true, force: true });
  }
});

test('writeAuditEvent preserves device id across writes', async () => {
  const auditRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'macroflow-audit-id-'));

  try {
    const first = await writeAuditEvent(
      {
        action: 'vba.run',
        channel: 'vba:run',
        outcome: 'blocked',
        reasonCode: 'POLICY_DENIED',
        workbook: { name: 'Book1.xlsm', path: 'C:\\Book1.xlsm' },
        target: { moduleName: '', macroName: 'MacroFlowModule1.RunA' },
        payload: { codeLength: 0 },
        result: { success: false, message: 'Denied.' }
      },
      { auditRoot }
    );

    const second = await writeAuditEvent(
      {
        action: 'vba.run',
        channel: 'vba:run',
        outcome: 'succeeded',
        reasonCode: 'OK',
        workbook: { name: 'Book1.xlsm', path: 'C:\\Book1.xlsm' },
        target: { moduleName: '', macroName: 'MacroFlowModule1.RunA' },
        payload: { codeLength: 0 },
        result: { success: true, message: 'Executed.' }
      },
      { auditRoot }
    );

    const rows = fs.readFileSync(first.filePath, 'utf8')
      .split(/\r?\n/)
      .filter(Boolean)
      .map((line) => JSON.parse(line));

    assert.equal(rows.length >= 2, true);
    assert.equal(rows[0].actor.deviceId, rows[1].actor.deviceId);
    assert.equal(second.success, true);
  } finally {
    fs.rmSync(auditRoot, { recursive: true, force: true });
  }
});
