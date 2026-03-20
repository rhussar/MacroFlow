const { autoUpdater, ipcMain, app } = require('electron');
const logger = require('./logger');

const UPDATE_FEED_URL = process.env.MACROFLOW_UPDATE_URL
  || 'https://pub-a7aa338dce944ce383fc182f58a87366.r2.dev/updates';

const CHECK_INTERVAL_MS = 4 * 60 * 60 * 1000;
const INITIAL_CHECK_DELAY_MS = 15_000;
const STARTUP_SETTLED_FALLBACK_MS = 45_000;

let mainWindowRef = null;
let checkTimer = null;
let initialTimer = null;
let startupFallbackTimer = null;
let updateDownloaded = false;
let pendingVersion = '';
let ipcRegistered = false;
let feedConfigured = false;
let updaterEventsRegistered = false;
let backgroundChecksStarted = false;
let startupSettled = false;

function sendToRenderer(channel, payload) {
  if (
    mainWindowRef &&
    !mainWindowRef.isDestroyed() &&
    mainWindowRef.webContents &&
    !mainWindowRef.webContents.isDestroyed()
  ) {
    mainWindowRef.webContents.send(channel, payload);
  }
}

function getUpdaterAvailability() {
  if (process.platform !== 'win32') {
    return { available: false, reason: 'not win32' };
  }

  if (process.env.NODE_ENV === 'development') {
    return { available: false, reason: 'development mode' };
  }

  if (!UPDATE_FEED_URL) {
    return { available: false, reason: 'no feed URL configured' };
  }

  return { available: true, reason: '' };
}

function clearTimeoutTimer(name) {
  if (name === 'initial' && initialTimer) {
    clearTimeout(initialTimer);
    initialTimer = null;
  }

  if (name === 'startupFallback' && startupFallbackTimer) {
    clearTimeout(startupFallbackTimer);
    startupFallbackTimer = null;
  }
}

function clearIntervalTimer() {
  if (!checkTimer) {
    return;
  }

  clearInterval(checkTimer);
  checkTimer = null;
}

function ensureFeedConfigured() {
  if (feedConfigured) {
    return true;
  }

  try {
    autoUpdater.setFeedURL({ url: UPDATE_FEED_URL });
    feedConfigured = true;
    logger.info('[AutoUpdater] initialized', {
      feedUrl: UPDATE_FEED_URL,
      currentVersion: app.getVersion(),
    });
    return true;
  } catch (err) {
    logger.error('[AutoUpdater] failed to set feed URL', { error: err.message });
    return false;
  }
}

function ensureUpdaterEventsRegistered() {
  if (updaterEventsRegistered) {
    return;
  }

  updaterEventsRegistered = true;

  autoUpdater.on('checking-for-update', () => {
    logger.info('[AutoUpdater] checking for update');
  });

  autoUpdater.on('update-available', () => {
    logger.info('[AutoUpdater] update available - downloading');
    sendToRenderer('updater:status', { status: 'downloading' });
  });

  autoUpdater.on('update-not-available', () => {
    logger.debug('[AutoUpdater] up to date');
  });

  autoUpdater.on('update-downloaded', (_event, _releaseNotes, releaseName) => {
    updateDownloaded = true;
    pendingVersion = releaseName || '';
    logger.info('[AutoUpdater] update downloaded', { releaseName });
    sendToRenderer('updater:status', {
      status: 'ready',
      version: pendingVersion,
    });
  });

  autoUpdater.on('error', (error) => {
    const message = error?.message || String(error);
    logger.error('[AutoUpdater] error', { error: message });
  });
}

function scheduleRecurringChecks() {
  if (checkTimer) {
    return;
  }

  checkTimer = setInterval(checkForUpdates, CHECK_INTERVAL_MS);
  if (typeof checkTimer.unref === 'function') {
    checkTimer.unref();
  }
}

function checkForUpdates() {
  try {
    autoUpdater.checkForUpdates();
  } catch (err) {
    logger.error('[AutoUpdater] checkForUpdates failed', { error: err.message });
  }
}

function startBackgroundChecks({ trigger = 'startup-settled', immediate = false } = {}) {
  const availability = getUpdaterAvailability();
  if (!availability.available) {
    logger.info('[AutoUpdater] background checks skipped', { reason: availability.reason });
    return false;
  }

  if (!ensureFeedConfigured()) {
    return false;
  }

  ensureUpdaterEventsRegistered();
  clearTimeoutTimer('startupFallback');

  if (backgroundChecksStarted) {
    if (immediate) {
      clearTimeoutTimer('initial');
      checkForUpdates();
      scheduleRecurringChecks();
    }
    return true;
  }

  backgroundChecksStarted = true;
  logger.info('[AutoUpdater] background checks armed', {
    trigger,
    initialDelayMs: immediate ? 0 : INITIAL_CHECK_DELAY_MS,
    intervalMs: CHECK_INTERVAL_MS,
  });

  if (immediate) {
    checkForUpdates();
    scheduleRecurringChecks();
    return true;
  }

  initialTimer = setTimeout(() => {
    initialTimer = null;
    checkForUpdates();
    scheduleRecurringChecks();
  }, INITIAL_CHECK_DELAY_MS);

  if (typeof initialTimer.unref === 'function') {
    initialTimer.unref();
  }

  return true;
}

function markStartupSettled() {
  startupSettled = true;
  if (backgroundChecksStarted) {
    return true;
  }

  logger.info('[AutoUpdater] startup settled; starting background checks');
  return startBackgroundChecks({ trigger: 'startup-settled' });
}

function scheduleStartupFallback() {
  if (startupFallbackTimer || backgroundChecksStarted) {
    return;
  }

  startupFallbackTimer = setTimeout(() => {
    startupFallbackTimer = null;
    if (backgroundChecksStarted) {
      return;
    }

    logger.info('[AutoUpdater] startup-settled fallback elapsed; starting background checks');
    startBackgroundChecks({ trigger: 'startup-fallback' });
  }, STARTUP_SETTLED_FALLBACK_MS);

  if (typeof startupFallbackTimer.unref === 'function') {
    startupFallbackTimer.unref();
  }
}

function initAutoUpdater(mainWindow) {
  // Register IPC handlers regardless of availability so the renderer
  // never hits "No handler registered" errors (e.g. in dev mode).
  if (!ipcRegistered) {
    ipcRegistered = true;

    ipcMain.on('updater:quit-and-install', () => {
      if (updateDownloaded) {
        logger.info('[AutoUpdater] quit-and-install requested');
        autoUpdater.quitAndInstall();
      }
    });

    ipcMain.handle('updater:check-now', async () => {
      const nextAvailability = getUpdaterAvailability();
      if (!nextAvailability.available) {
        return { success: false, reason: nextAvailability.reason };
      }

      startBackgroundChecks({ trigger: 'manual-check', immediate: true });
      return { success: true };
    });

    ipcMain.handle('updater:status', () => {
      return {
        updateDownloaded,
        pendingVersion,
        currentVersion: app.getVersion(),
        feedUrl: UPDATE_FEED_URL || null,
        backgroundChecksStarted,
        startupSettled,
      };
    });

    ipcMain.on('app:startup-settled', () => {
      markStartupSettled();
    });
  }

  const availability = getUpdaterAvailability();
  if (!availability.available) {
    if (availability.reason === 'no feed URL configured') {
      logger.warn(
        '[AutoUpdater] skipped - MACROFLOW_UPDATE_URL not set. ' +
        'Set this env var to enable auto-updates.'
      );
      return;
    }

    logger.info('[AutoUpdater] skipped - ' + availability.reason);
    return;
  }

  mainWindowRef = mainWindow;

  logger.info('[AutoUpdater] deferring background checks until startup settled', {
    fallbackMs: STARTUP_SETTLED_FALLBACK_MS,
    initialDelayMs: INITIAL_CHECK_DELAY_MS,
  });
  scheduleStartupFallback();
}

function stopAutoUpdater() {
  clearTimeoutTimer('initial');
  clearTimeoutTimer('startupFallback');
  clearIntervalTimer();
  mainWindowRef = null;
}

module.exports = { initAutoUpdater, stopAutoUpdater };
