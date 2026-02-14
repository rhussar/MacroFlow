let baseLogger;

try {
  baseLogger = require('electron-log/main');
  if (typeof baseLogger.initialize === 'function') {
    baseLogger.initialize();
  }
} catch {
  baseLogger = require('electron-log');
}

if (baseLogger?.transports?.file) {
  baseLogger.transports.file.level = 'info';
  baseLogger.transports.file.format = '[{y}-{m}-{d} {h}:{i}:{s}.{ms}] [{level}] {text}';
}

if (baseLogger?.transports?.console) {
  baseLogger.transports.console.level = process.env.NODE_ENV === 'development' ? 'debug' : 'info';
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
