/**
 * Diagnostics Module - Lightweight logging and diagnostic utilities
 * 
 * Provides consistent logging across the application and diagnostic
 * functions for troubleshooting Excel integration issues.
 */

const fs = require('fs-extra');
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');
const persistentLogger = require('./logger');

const execAsync = promisify(exec);

// ============================================================================
// LOGGING
// ============================================================================

const LOG_LEVELS = {
    DEBUG: 0,
    INFO: 1,
    WARN: 2,
    ERROR: 3
};

// Current log level (can be set via environment variable)
let currentLogLevel = LOG_LEVELS[process.env.MACROFLOW_LOG_LEVEL?.toUpperCase()] ?? LOG_LEVELS.INFO;

// In-memory log buffer for recent entries (circular buffer)
const LOG_BUFFER_SIZE = 500;
const logBuffer = [];
let logBufferIndex = 0;

/**
 * Format a timestamp for logging
 */
function getTimestamp() {
    return new Date().toISOString();
}

/**
 * Add an entry to the log buffer
 */
function addToBuffer(level, category, message, details) {
    const entry = {
        timestamp: getTimestamp(),
        level,
        category,
        message,
        details: details || null
    };

    if (logBuffer.length < LOG_BUFFER_SIZE) {
        logBuffer.push(entry);
    } else {
        logBuffer[logBufferIndex] = entry;
        logBufferIndex = (logBufferIndex + 1) % LOG_BUFFER_SIZE;
    }
}

/**
 * Core logging function
 */
function log(level, category, message, details = null) {
    if (LOG_LEVELS[level] < currentLogLevel) {
        return;
    }

    addToBuffer(level, category, message, details);

    const timestamp = getTimestamp().substring(11, 23); // HH:mm:ss.sss
    const detailStr = details ? ` ${JSON.stringify(details)}` : '';
    const prefix = `[${level}][${category}][${timestamp}]`;

    const logFn = level === 'ERROR' ? console.error :
        level === 'WARN' ? console.warn :
            console.log;

    logFn(`${prefix} ${message}${detailStr}`);

    // Mirror diagnostics logs to persistent electron-log file so shutdown traces
    // are still available after the app process exits.
    const normalizedLevel = String(level || '').toLowerCase();
    const fileMessage = `[${category}] ${message}`;
    if (typeof persistentLogger[normalizedLevel] === 'function') {
        persistentLogger[normalizedLevel](fileMessage, details || undefined);
    } else {
        persistentLogger.info(fileMessage, details || undefined);
    }
}

// Convenience logging methods
const logger = {
    debug: (category, message, details) => log('DEBUG', category, message, details),
    info: (category, message, details) => log('INFO', category, message, details),
    warn: (category, message, details) => log('WARN', category, message, details),
    error: (category, message, details) => log('ERROR', category, message, details),

    /**
     * Set the minimum log level
     */
    setLevel: (level) => {
        if (LOG_LEVELS[level] !== undefined) {
            currentLogLevel = LOG_LEVELS[level];
            logger.info('Logger', `Log level set to ${level}`);
        }
    },

    /**
     * Get recent log entries
     */
    getRecentLogs: (count = 100) => {
        const entries = [];
        const start = logBuffer.length < LOG_BUFFER_SIZE ? 0 : logBufferIndex;
        const total = Math.min(count, logBuffer.length);

        for (let i = 0; i < total; i++) {
            const idx = (start + logBuffer.length - 1 - i) % logBuffer.length;
            if (logBuffer[idx]) {
                entries.push(logBuffer[idx]);
            }
        }

        return entries;
    }
};

// ============================================================================
// EXCEL DIAGNOSTICS
// ============================================================================

/**
 * Check if Excel is currently running
 */
async function isExcelRunning() {
    try {
        const { stdout } = await execAsync('tasklist /FI "IMAGENAME eq EXCEL.EXE" /FO CSV /NH');
        return stdout.toLowerCase().includes('excel.exe');
    } catch (error) {
        logger.error('Diagnostics', 'Failed to check if Excel is running', { error: error.message });
        return false;
    }
}

/**
 * Get Excel process information
 */
async function getExcelProcessInfo() {
    try {
        const { stdout } = await execAsync('tasklist /FI "IMAGENAME eq EXCEL.EXE" /FO CSV');
        const lines = stdout.trim().split('\n');
        if (lines.length < 2) {
            return { running: false, processes: [] };
        }

        const processes = [];
        for (let i = 1; i < lines.length; i++) {
            const parts = lines[i].split(',').map(p => p.replace(/"/g, ''));
            if (parts.length >= 5) {
                processes.push({
                    name: parts[0],
                    pid: parseInt(parts[1], 10),
                    memoryUsage: parts[4]
                });
            }
        }

        return { running: processes.length > 0, processes };
    } catch (error) {
        logger.error('Diagnostics', 'Failed to get Excel process info', { error: error.message });
        return { running: false, processes: [], error: error.message };
    }
}

/**
 * Check Excel modal dialog state (heuristic)
 * Returns true if Excel MAY be showing a modal dialog
 */
async function checkExcelModalState() {
    try {
        // This is a heuristic - we check if Excel's main window is responsive
        // by trying to enumerate its windows
        const script = `
      $ErrorActionPreference = 'SilentlyContinue'
      try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        if ($excel) {
          $interactive = $excel.Interactive
          $ready = $excel.Ready
          Write-Output "Interactive:$interactive,Ready:$ready"
        } else {
          Write-Output "NoExcel"
        }
      } catch {
        Write-Output "Error:$_"
      }
    `;

        const encoded = Buffer.from(script, 'utf16le').toString('base64');
        const { stdout } = await execAsync(`powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ${encoded}`);

        const result = stdout.trim();
        if (result === 'NoExcel') {
            return { hasModal: false, reason: 'excel_not_found' };
        }
        if (result.startsWith('Error:')) {
            return { hasModal: true, reason: 'excel_possibly_blocked', error: result };
        }

        const match = result.match(/Interactive:(True|False),Ready:(True|False)/i);
        if (match) {
            const interactive = match[1].toLowerCase() === 'true';
            const ready = match[2].toLowerCase() === 'true';
            const hasModal = !interactive || !ready;
            return {
                hasModal,
                interactive,
                ready,
                reason: hasModal ? 'excel_not_ready' : 'excel_ready'
            };
        }

        return { hasModal: false, reason: 'unknown' };
    } catch (error) {
        // If we can't query Excel, assume it might be showing a modal
        logger.warn('Diagnostics', 'Could not check Excel modal state', { error: error.message });
        return { hasModal: true, reason: 'check_failed', error: error.message };
    }
}

// ============================================================================
// ADD-IN DIAGNOSTICS
// ============================================================================

const ADDIN_FILE_NAME = 'MacroFlow.xlam';

/**
 * Check add-in installation status
 */
async function checkAddinStatus() {
    const result = {
        installed: false,
        registered: false,
        loadedInExcel: false,
        filePath: null,
        installLocation: null,
        errors: []
    };

    // Check if file exists in XLSTART (current install location).
    // Fallback to legacy AddIns location for backward compatibility diagnostics.
    const appData = process.env.APPDATA;
    if (appData) {
        const xlstartPath = path.join(appData, 'Microsoft', 'Excel', 'XLSTART', ADDIN_FILE_NAME);
        const legacyAddinsPath = path.join(appData, 'Microsoft', 'AddIns', ADDIN_FILE_NAME);
        try {
            const xlstartExists = await fs.pathExists(xlstartPath);
            const legacyExists = xlstartExists ? false : await fs.pathExists(legacyAddinsPath);
            const resolvedPath = xlstartExists ? xlstartPath : (legacyExists ? legacyAddinsPath : '');

            result.installed = Boolean(resolvedPath);
            result.registered = result.installed;

            if (resolvedPath) {
                result.filePath = resolvedPath;
                result.installLocation = xlstartExists ? 'xlstart' : 'legacy_addins';
                const stat = await fs.stat(resolvedPath);
                result.fileSize = stat.size;
                result.fileModified = stat.mtime.toISOString();
            }
        } catch (error) {
            result.errors.push(`File check failed: ${error.message}`);
        }
    }

    // Check if add-in is loaded in running Excel
    try {
        const script = `
      $ErrorActionPreference = 'SilentlyContinue'
      try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        if ($excel) {
          $addins = $excel.AddIns
          foreach ($addin in $addins) {
            if ($addin.Name -like '*MacroFlow*' -and $addin.Installed) {
              Write-Output "LOADED:$($addin.Name):$($addin.FullName)"
              exit
            }
          }
          Write-Output "NOT_LOADED"
        } else {
          Write-Output "NO_EXCEL"
        }
      } catch {
        Write-Output "ERROR:$_"
      }
    `;

        const encoded = Buffer.from(script, 'utf16le').toString('base64');
        const { stdout } = await execAsync(`powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ${encoded}`);

        const output = stdout.trim();
        if (output.startsWith('LOADED:')) {
            result.loadedInExcel = true;
            const parts = output.split(':');
            result.loadedAddinName = parts[1];
            result.loadedAddinPath = parts[2];
        } else if (output === 'NO_EXCEL') {
            result.excelRunning = false;
        } else if (output.startsWith('ERROR:')) {
            result.errors.push(`Excel query error: ${output}`);
        }
    } catch (error) {
        result.errors.push(`Live check failed: ${error.message}`);
    }

    return result;
}

/**
 * Check ribbon consistency by verifying the add-in's RibbonX
 */
async function checkRibbonStatus() {
    const result = {
        hasRibbon: false,
        ribbonErrors: [],
        ribbonInfo: null
    };

    try {
        // Check if the add-in file has a customUI folder (indicates ribbon customization)
        const appData = process.env.APPDATA;
        if (!appData) {
            result.ribbonErrors.push('APPDATA not found');
            return result;
        }

        const xlstartPath = path.join(appData, 'Microsoft', 'Excel', 'XLSTART', ADDIN_FILE_NAME);
        const legacyAddinsPath = path.join(appData, 'Microsoft', 'AddIns', ADDIN_FILE_NAME);
        let addinPath = xlstartPath;

        if (!await fs.pathExists(addinPath)) {
            if (await fs.pathExists(legacyAddinsPath)) {
                addinPath = legacyAddinsPath;
            } else {
                result.ribbonErrors.push('Add-in file not found');
                return result;
            }
        }

        if (!await fs.pathExists(addinPath)) {
            result.ribbonErrors.push('Add-in file not found');
            return result;
        }

        // The xlam file is a zip - we could inspect it, but for now just check if it loads
        result.hasRibbon = true; // Assume ribbon exists if add-in exists
        result.ribbonInfo = {
            path: addinPath,
            note: 'Ribbon is defined in the .xlam file. If not visible, restart Excel and verify MacroFlow.xlam exists in XLSTART.'
        };

    } catch (error) {
        result.ribbonErrors.push(`Ribbon check failed: ${error.message}`);
    }

    return result;
}

// ============================================================================
// SYSTEM DIAGNOSTICS
// ============================================================================

/**
 * Collect system diagnostic information
 */
async function collectDiagnostics() {
    logger.info('Diagnostics', 'Collecting system diagnostics...');

    const diagnostics = {
        timestamp: getTimestamp(),
        system: {
            platform: process.platform,
            arch: process.arch,
            nodeVersion: process.version,
            electronVersion: process.versions.electron,
            env: {
                NODE_ENV: process.env.NODE_ENV,
                APPDATA: process.env.APPDATA ? '[SET]' : '[NOT SET]'
            }
        },
        excel: await getExcelProcessInfo(),
        excelModal: await checkExcelModalState(),
        addin: await checkAddinStatus(),
        ribbon: await checkRibbonStatus(),
        recentLogs: logger.getRecentLogs(50)
    };

    logger.info('Diagnostics', 'Diagnostics collected', {
        excelRunning: diagnostics.excel.running,
        addinInstalled: diagnostics.addin.installed,
        addinLoaded: diagnostics.addin.loadedInExcel
    });

    return diagnostics;
}

/**
 * Write diagnostics to a file (for support/debugging)
 */
async function writeDiagnosticsToFile(outputPath) {
    const diagnostics = await collectDiagnostics();
    const content = JSON.stringify(diagnostics, null, 2);

    await fs.writeFile(outputPath, content, 'utf8');
    logger.info('Diagnostics', `Diagnostics written to ${outputPath}`);

    return outputPath;
}

module.exports = {
    logger,
    isExcelRunning,
    getExcelProcessInfo,
    checkExcelModalState,
    checkAddinStatus,
    checkRibbonStatus,
    collectDiagnostics,
    writeDiagnosticsToFile
};
