const fs = require('fs-extra');
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');

const execAsync = promisify(exec);

const ADDIN_FILE_NAME = 'MacroFlow.xlam';

// Lazy-loaded logger to avoid circular dependencies
let _logger = null;
function getLogger() {
  if (!_logger) {
    try {
      _logger = require('./diagnostics').logger;
    } catch {
      // Fallback to console if diagnostics not available
      _logger = {
        debug: (cat, msg, details) => console.log(`[DEBUG][${cat}] ${msg}`, details || ''),
        info: (cat, msg, details) => console.log(`[INFO][${cat}] ${msg}`, details || ''),
        warn: (cat, msg, details) => console.warn(`[WARN][${cat}] ${msg}`, details || ''),
        error: (cat, msg, details) => console.error(`[ERROR][${cat}] ${msg}`, details || '')
      };
    }
  }
  return _logger;
}

function log(level, message, details = null) {
  const logger = getLogger();
  logger[level]('AddinInstaller', message, details);
}

async function isFileLocked(filePath) {
  if (!await fs.pathExists(filePath)) return false;
  try {
    const fd = await fs.open(filePath, 'r+');
    await fd.close();
    return false;
  } catch {
    return true;
  }
}

async function isExcelRunning() {
  try {
    const { stdout } = await execAsync('tasklist /FI "IMAGENAME eq EXCEL.EXE"');
    return stdout.toUpperCase().includes('EXCEL.EXE');
  } catch {
    return false;
  }
}

// ============================================================================
// PATH HELPERS
// ============================================================================

/**
 * Get the path to the MacroFlow.xlam file
 * Handles both development and production environments
 */
function getAddinSourcePath() {
  const isDev = process.env.NODE_ENV === 'development';

  if (isDev) {
    // Development: resources folder is at project root
    return path.join(__dirname, '../Resources', ADDIN_FILE_NAME);
  } else {
    // Production: extraResource copies to process.resourcesPath/resources
    if (process.resourcesPath) {
      const candidates = [
        // Some packagers copy individual files directly into the resources root.
        path.join(process.resourcesPath, ADDIN_FILE_NAME),
        path.join(process.resourcesPath, 'Resources', ADDIN_FILE_NAME),
        path.join(process.resourcesPath, 'resources', ADDIN_FILE_NAME),
      ];
      for (const candidate of candidates) {
        if (fs.existsSync(candidate)) {
          return candidate;
        }
      }
    }
    // Fallback
    return path.join(__dirname, '../Resources', ADDIN_FILE_NAME);
  }
}

/**
 * Get the Excel XLSTART folder path
 * Excel auto-loads files placed here
 */
function getXlstartFolder() {
  const appData = process.env.APPDATA;
  if (!appData) {
    throw new Error('APPDATA environment variable not found');
  }
  return path.join(appData, 'Microsoft', 'Excel', 'XLSTART');
}

// ============================================================================
// FILE OPERATIONS
// ============================================================================

/**
 * Check if two files are identical by comparing their content
 */
async function areFilesIdentical(file1, file2) {
  try {
    const [content1, content2] = await Promise.all([
      fs.readFile(file1),
      fs.readFile(file2)
    ]);
    return content1.equals(content2);
  } catch (error) {
    return false;
  }
}

/**
 * Copy the add-in file to the XLSTART folder
 */
async function copyAddinToXlstartFolder(sourcePath) {
  const xlstartFolder = getXlstartFolder();
  const destPath = path.join(xlstartFolder, ADDIN_FILE_NAME);

  // Ensure XLSTART directory exists
  await fs.ensureDir(xlstartFolder);

  // Check if already installed and up-to-date
  if (await fs.pathExists(destPath)) {
    const identical = await areFilesIdentical(sourcePath, destPath);
    if (identical) {
      return { copied: false, path: destPath, reason: 'already_up_to_date' };
    }
  }

  // Copy the file
  await fs.copy(sourcePath, destPath, { overwrite: true });
  return { copied: true, path: destPath };
}

async function cleanupLegacyInstalls() {
  try {
    const appData = process.env.APPDATA;
    if (appData) {
      const legacyAddinsPath = path.join(appData, 'Microsoft', 'AddIns', ADDIN_FILE_NAME);
      if (await fs.pathExists(legacyAddinsPath)) {
        await fs.remove(legacyAddinsPath);
        log('info', 'Removed legacy AddIns file', { path: legacyAddinsPath });
      }
    }
  } catch (error) {
    log('debug', 'Legacy AddIns cleanup failed (ignored)', { error: error.message });
  }
}

// ============================================================================
// MAIN INSTALLER FUNCTION
// ============================================================================

/**
 * Install the Excel Add-in:
 * 1. Copy to %APPDATA%\Microsoft\Excel\XLSTART
 * 2. Attempt to load into running Excel instance
 */
async function installExcelAddin() {
  log('info', 'Starting Excel Add-in installation...');

  if (process.platform !== 'win32') {
    log('warn', 'Installation skipped - Windows only');
    return { success: false, error: 'Windows only.' };
  }

  try {
    await cleanupLegacyInstalls();

    // Step 1: Get source path
    const sourcePath = getAddinSourcePath();
    log('debug', 'Add-in source path resolved', { path: sourcePath });

    if (!await fs.pathExists(sourcePath)) {
      const error = `${ADDIN_FILE_NAME} not found at ${sourcePath}`;
      log('error', error);
      throw new Error(error);
    }

    // Get file info for logging
    const sourceStats = await fs.stat(sourcePath);
    log('info', 'Add-in source file found', {
      size: sourceStats.size,
      modified: sourceStats.mtime.toISOString()
    });

    // Step 2: Copy to XLSTART folder
    const xlstartFolder = getXlstartFolder();
    const destinationPath = path.join(xlstartFolder, ADDIN_FILE_NAME);
    if (await isFileLocked(destinationPath)) {
      const error = 'Please close Excel and try again.';
      log('warn', 'Destination file is locked', { path: destinationPath });
      return { success: false, error, needsExcelClosed: true };
    }

    log('debug', 'Copying add-in to XLSTART folder...');
    const copyResult = await copyAddinToXlstartFolder(sourcePath);
    log('info', copyResult.copied ? 'Add-in file copied' : 'Add-in file already up-to-date', {
      path: copyResult.path,
      reason: copyResult.reason
    });

    const excelRunning = await isExcelRunning();

    const result = {
      success: true,
      addinPath: copyResult.path,
      fileCopied: copyResult.copied,
      restartRequired: excelRunning,
      message: excelRunning ? 'Restart Excel to activate MacroFlow.' : undefined
    };

    log('info', 'Excel Add-in installation completed successfully', {
      fileCopied: result.fileCopied,
      restartRequired: result.restartRequired
    });

    return result;

  } catch (error) {
    log('error', 'Installation failed', { error: error.message, stack: error.stack });
    return {
      success: false,
      error: error.message
    };
  }
}


/**
 * Uninstall the Excel Add-in:
 * 1. Delete from XLSTART folder
 */
async function uninstallExcelAddin() {
  log('info', 'Starting Excel Add-in uninstallation...');

  if (process.platform !== 'win32') {
    log('warn', 'Uninstallation skipped - Windows only');
    return { success: false, error: 'Windows only.' };
  }

  try {
    await cleanupLegacyInstalls();

    const xlstartFolder = getXlstartFolder();
    const addinPath = path.join(xlstartFolder, ADDIN_FILE_NAME);

    // Delete file
    if (await fs.pathExists(addinPath)) {
      await fs.remove(addinPath);
      log('info', 'Add-in file deleted from XLSTART', { path: addinPath });
    } else {
      log('debug', 'Add-in file not found in XLSTART (already deleted or never installed)');
    }

    log('info', 'Excel Add-in uninstallation completed successfully');
    return { success: true };
  } catch (error) {
    log('error', 'Uninstallation failed', { error: error.message });
    return { success: false, error: error.message };
  }
}

module.exports = { installExcelAddin, uninstallExcelAddin };
