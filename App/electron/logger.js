let baseLogger;

try {
  baseLogger = require('electron-log/main');
  if (typeof baseLogger.initialize === 'function') {
    baseLogger.initialize();
  }
} catch {
  baseLogger = require('electron-log');
}

const LOG_LEVELS = new Set(['error', 'warn', 'info', 'verbose', 'debug', 'silly']);
const envLevel = String(process.env.MACROFLOW_LOG_LEVEL || '').trim().toLowerCase();
const resolvedLevel = LOG_LEVELS.has(envLevel)
  ? envLevel
  : (process.env.NODE_ENV === 'development' ? 'debug' : 'info');

if (baseLogger?.transports?.file) {
  baseLogger.transports.file.level = resolvedLevel;
  baseLogger.transports.file.format = '[{y}-{m}-{d} {h}:{i}:{s}.{ms}] [{level}] {text}';
}

if (baseLogger?.transports?.console) {
  baseLogger.transports.console.level = resolvedLevel;
}

function stringifyMeta(meta) {
  if (!meta || typeof meta !== 'object' || Array.isArray(meta)) {
    return '';
  }
  try {
    return ` ${JSON.stringify(meta)}`;
  } catch {
    return '';
  }
}

function log(level, message, meta) {
  const text = `${message}${stringifyMeta(meta)}`;
  if (typeof baseLogger[level] === 'function') {
    baseLogger[level](text);
    return;
  }
  baseLogger.info(text);
}

module.exports = {
  debug(message, meta) {
    log('debug', message, meta);
  },
  info(message, meta) {
    log('info', message, meta);
  },
  warn(message, meta) {
    log('warn', message, meta);
  },
  error(message, meta) {
    log('error', message, meta);
  }
};
