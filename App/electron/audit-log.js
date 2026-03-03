const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

const RETENTION_DAYS = 90;
const MS_PER_DAY = 24 * 60 * 60 * 1000;
const DEVICE_ID_FILE_NAME = 'device-id';
const AUDIT_DIRECTORY_NAME = 'audit';

function toSafeString(value) {
  return String(value || '').trim();
}

function normalizeLineEndings(value) {
  return String(value || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function toSha256Hex(text) {
  return crypto.createHash('sha256').update(String(text || ''), 'utf8').digest('hex');
}

function getAuditRoot(options = {}) {
  const fromOptions = toSafeString(options.auditRoot);
  if (fromOptions) {
    return path.resolve(fromOptions);
  }

  const appData = toSafeString(process.env.APPDATA);
  if (appData) {
    return path.join(appData, 'MacroFlow');
  }

  return path.join(process.cwd(), '.macroflow');
}

function ensureDirectory(directoryPath) {
  fs.mkdirSync(directoryPath, { recursive: true });
}

function pruneAuditFiles(auditDirectoryPath, nowMs) {
  const cutoffMs = nowMs - (RETENTION_DAYS * MS_PER_DAY);
  let entries = [];
  try {
    entries = fs.readdirSync(auditDirectoryPath, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    if (!entry.isFile()) {
      continue;
    }
    if (!/^audit-\d{4}-\d{2}-\d{2}\.jsonl$/i.test(entry.name)) {
      continue;
    }

    const fullPath = path.join(auditDirectoryPath, entry.name);
    try {
      const stats = fs.statSync(fullPath);
      if (Number(stats.mtimeMs) < cutoffMs) {
        fs.unlinkSync(fullPath);
      }
    } catch {
      // Ignore retention cleanup failures.
    }
  }
}

function getDeviceId(rootPath) {
  const deviceIdPath = path.join(rootPath, DEVICE_ID_FILE_NAME);
  try {
    const existing = toSafeString(fs.readFileSync(deviceIdPath, 'utf8'));
    if (existing) {
      return existing;
    }
  } catch {
    // Fall through to create a new device id.
  }

  const nextId = crypto.randomUUID();
  try {
    fs.writeFileSync(deviceIdPath, nextId, 'utf8');
  } catch {
    // If write fails, still return generated id for this process.
  }
  return nextId;
}

function normalizeWorkbookIdentifier(workbook = {}) {
  const workbookName = toSafeString(workbook.name);
  const workbookPath = toSafeString(workbook.path);
  const identifierSource = `${workbookPath.toLowerCase()}::${workbookName.toLowerCase()}`;
  return {
    name: workbookName,
    path: workbookPath,
    identifierHash: toSha256Hex(identifierSource)
  };
}

function toDateStamp(now) {
  const year = String(now.getFullYear());
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function buildAuditRecord(entry = {}, options = {}) {
  const now = options.now instanceof Date ? options.now : new Date();
  const codeText = entry?.payload && Object.prototype.hasOwnProperty.call(entry.payload, 'code')
    ? normalizeLineEndings(entry.payload.code)
    : '';
  const codeLength = codeText.length;

  return {
    eventId: crypto.randomUUID(),
    timestamp: now.toISOString(),
    action: toSafeString(entry.action),
    channel: toSafeString(entry.channel),
    outcome: toSafeString(entry.outcome),
    reasonCode: toSafeString(entry.reasonCode),
    actor: {
      username: toSafeString(entry?.actor?.username) || toSafeString(process.env.USERNAME) || toSafeString(process.env.USER) || 'unknown',
      deviceId: toSafeString(entry?.actor?.deviceId) || toSafeString(options.deviceId)
    },
    workbook: normalizeWorkbookIdentifier(entry.workbook),
    target: {
      moduleName: toSafeString(entry?.target?.moduleName),
      macroName: toSafeString(entry?.target?.macroName)
    },
    payload: {
      codeHash: codeLength > 0 ? toSha256Hex(codeText) : null,
      codeLength: Number.isFinite(Number(entry?.payload?.codeLength))
        ? Number(entry.payload.codeLength)
        : codeLength
    },
    result: {
      success: Boolean(entry?.result?.success),
      message: toSafeString(entry?.result?.message)
    }
  };
}

async function writeAuditEvent(entry = {}, options = {}) {
  const rootPath = getAuditRoot(options);
  const auditDirectoryPath = path.join(rootPath, AUDIT_DIRECTORY_NAME);
  ensureDirectory(auditDirectoryPath);

  const deviceId = getDeviceId(rootPath);
  const now = options.now instanceof Date ? options.now : new Date();
  const nowMs = now.getTime();
  pruneAuditFiles(auditDirectoryPath, nowMs);

  const record = buildAuditRecord(entry, { now, deviceId });
  const fileName = `audit-${toDateStamp(now)}.jsonl`;
  const filePath = path.join(auditDirectoryPath, fileName);
  await fs.promises.appendFile(filePath, `${JSON.stringify(record)}\n`, 'utf8');

  return {
    success: true,
    filePath,
    record
  };
}

function initializeAuditLog(options = {}) {
  const rootPath = getAuditRoot(options);
  const auditDirectoryPath = path.join(rootPath, AUDIT_DIRECTORY_NAME);
  ensureDirectory(auditDirectoryPath);

  const now = options.now instanceof Date ? options.now : new Date();
  pruneAuditFiles(auditDirectoryPath, now.getTime());

  return {
    success: true,
    rootPath,
    auditDirectoryPath
  };
}

module.exports = {
  initializeAuditLog,
  writeAuditEvent,
  buildAuditRecord,
  toSha256Hex
};
