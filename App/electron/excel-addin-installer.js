const fs = require('fs-extra');
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');

const execAsync = promisify(exec);

const ADDIN_FILE_NAME = 'MacroFlow.xlam';
const DEFAULT_EXCEL_OPTIONS_KEY = 'HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Options';

function escapePowerShellSingleQuotedString(value) {
  return String(value).replace(/'/g, "''");
}

function buildPowerShellEncodedCommand(script) {
  return Buffer.from(String(script), 'utf16le').toString('base64');
}

async function runPowerShellEncoded(script) {
  const encoded = buildPowerShellEncodedCommand(script);
  return execAsync(`powershell.exe -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -EncodedCommand ${encoded}`);
}

async function doesRegistryKeyExist(key) {
  try {
    await execAsync(`reg query "${key}" 2>nul`);
    return true;
  } catch {
    return false;
  }
}

async function getInstalledExcelMajorVersion() {
  const versions = ['16.0', '15.0', '14.0'];
  const installRootTemplates = [
    'HKLM\\Software\\Microsoft\\Office\\{ver}\\Excel\\InstallRoot',
    'HKLM\\Software\\WOW6432Node\\Microsoft\\Office\\{ver}\\Excel\\InstallRoot',
  ];

  for (const ver of versions) {
    for (const template of installRootTemplates) {
      const key = template.replace('{ver}', ver);
      if (await doesRegistryKeyExist(key)) {
        return ver;
      }
    }
  }

  return null;
}

async function getExcelOptionsRegistryKey() {
  // Excel 365 / Office 2016+ is typically 16.0, but older installations can differ.
  const versions = ['16.0', '15.0', '14.0'];

  // Prefer whichever HKCU key exists (Excel has likely been run at least once)
  for (const v of versions) {
    const key = `HKCU\\Software\\Microsoft\\Office\\${v}\\Excel\\Options`;
    if (await doesRegistryKeyExist(key)) {
      return key;
    }
  }

  // If HKCU keys aren't present yet, infer from install roots in HKLM
  const installed = await getInstalledExcelMajorVersion();
  if (installed) {
    return `HKCU\\Software\\Microsoft\\Office\\${installed}\\Excel\\Options`;
  }

  // If nothing matches, we still write to the default (registry write will create it).
  return DEFAULT_EXCEL_OPTIONS_KEY;
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
 * Get the Microsoft AddIns folder path
 * This is where Excel expects registered add-ins to live
 */
function getAddInsFolder() {
  const appData = process.env.APPDATA;
  if (!appData) {
    throw new Error('APPDATA environment variable not found');
  }
  return path.join(appData, 'Microsoft', 'AddIns');
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
 * Copy the add-in file to the AddIns folder
 */
async function copyAddinToAddInsFolder(sourcePath) {
  const addInsFolder = getAddInsFolder();
  const destPath = path.join(addInsFolder, ADDIN_FILE_NAME);

  // Ensure AddIns directory exists
  await fs.ensureDir(addInsFolder);

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

// ============================================================================
// REGISTRY OPERATIONS
// ============================================================================

/**
 * Find the next available OPEN slot in Excel registry
 * Excel uses OPEN, OPEN1, OPEN2, etc. for auto-load add-ins
 */
async function findNextOpenSlot() {
  const baseKey = await getExcelOptionsRegistryKey();

  // Check OPEN first
  try {
    const { stdout } = await execAsync(`reg query "${baseKey}" /v OPEN 2>nul`);
    // OPEN exists, check if it's our add-in
    if (stdout.includes(ADDIN_FILE_NAME)) {
      return { slot: 'OPEN', alreadyRegistered: true };
    }
  } catch {
    // OPEN doesn't exist, we can use it
    return { slot: 'OPEN', alreadyRegistered: false };
  }

  // Check OPEN1 through OPEN10
  for (let i = 1; i <= 10; i++) {
    const valueName = `OPEN${i}`;
    try {
      const { stdout } = await execAsync(`reg query "${baseKey}" /v ${valueName} 2>nul`);
      if (stdout.includes(ADDIN_FILE_NAME)) {
        return { slot: valueName, alreadyRegistered: true };
      }
    } catch {
      // This slot is free
      return { slot: valueName, alreadyRegistered: false };
    }
  }

  throw new Error('All Excel Add-in slots (OPEN-OPEN10) are full.');
}

/**
 * Register the add-in in the Windows registry for Excel auto-load
 * @param {string} destinationPath - The path where the add-in was copied to (in AddIns folder)
 */
async function registerAddinInRegistry(destinationPath) {
  const baseKey = await getExcelOptionsRegistryKey();

  // Find available slot
  const { slot, alreadyRegistered } = await findNextOpenSlot();

  if (alreadyRegistered) {
    return { registered: false, slot, reason: 'already_registered' };
  }

  // Format: /R "C:\Users\...\AppData\Roaming\Microsoft\AddIns\MacroFlow.xlam"
  // The /R switch tells Excel to load as a hidden add-in (not a visible workbook)
  const valueData = `/R "${destinationPath}"`;

  const psKeyPath = baseKey.replace(/^HKCU\\/, 'HKCU:\\');
  const escapedKeyPath = escapePowerShellSingleQuotedString(psKeyPath);
  const escapedSlot = escapePowerShellSingleQuotedString(slot);
  const escapedValueData = escapePowerShellSingleQuotedString(valueData);

  try {
    await runPowerShellEncoded(
      [
        "$ErrorActionPreference = 'Stop'",
        `$keyPath = '${escapedKeyPath}'`,
        'New-Item -Path $keyPath -Force | Out-Null',
        `New-ItemProperty -Path $keyPath -Name '${escapedSlot}' -PropertyType String -Value '${escapedValueData}' -Force | Out-Null`,
      ].join('; ')
    );
    return { registered: true, slot, path: destinationPath };
  } catch (error) {
    throw new Error(`Registry write failed: ${error.message}`);
  }
}

async function tryLoadAddinIntoRunningExcel(destinationPath) {
  const escapedPath = escapePowerShellSingleQuotedString(destinationPath);
  const escapedName = escapePowerShellSingleQuotedString(ADDIN_FILE_NAME);

  const script = [
    "$ErrorActionPreference = 'Stop'",
    `$addinPath = '${escapedPath}'`,
    `$addinName = '${escapedName}'`,
    '$excel = $null',
    "try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { exit 2 }",
    '$addin = $excel.AddIns | Where-Object { $_.FullName -eq $addinPath -or $_.Name -eq $addinName } | Select-Object -First 1',
    'if (-not $addin) { $addin = $excel.AddIns.Add($addinPath, $false) }',
    '$addin.Installed = $true',
  ].join('; ');

  try {
    await runPowerShellEncoded(script);
    return { attempted: true, loaded: true };
  } catch (error) {
    // Exit code 2 means no running Excel instance to attach to (not an error for install)
    if (typeof error?.code === 'number' && error.code === 2) {
      return { attempted: true, loaded: false, reason: 'excel_not_running' };
    }
    return { attempted: true, loaded: false, reason: 'load_failed', error: error.message };
  }
}

// ============================================================================
// MAIN INSTALLER FUNCTION
// ============================================================================

/**
 * Install the Excel Add-in:
 * 1. Copy to %APPDATA%\Microsoft\AddIns
 * 2. Register in Windows Registry for auto-load
 */
async function installExcelAddin() {
  if (process.platform !== 'win32') {
    return { success: false, error: 'Windows only.' };
  }
  try {
    // Step 1: Get source path
    const sourcePath = getAddinSourcePath();

    if (!await fs.pathExists(sourcePath)) {
      throw new Error(`${ADDIN_FILE_NAME} not found at ${sourcePath}`);
    }

    // Step 2: Copy to AddIns folder
    const copyResult = await copyAddinToAddInsFolder(sourcePath);

    // Step 2.5: Clean up any old/stale MacroFlow registry entries so Excel doesn't try to load missing add-ins.
    const baseKey = await getExcelOptionsRegistryKey();
    const slots = ['OPEN', ...Array.from({ length: 10 }, (_, i) => `OPEN${i + 1}`)];
    for (const slot of slots) {
      try {
        const { stdout } = await execAsync(`reg query "${baseKey}" /v ${slot} 2>nul`);
        const hasCurrent = stdout.includes(ADDIN_FILE_NAME);
        if (hasCurrent) {
          // If it already points at the correct file path, keep it.
          if (stdout.includes(copyResult.path)) {
            continue;
          }
          await execAsync(`reg delete "${baseKey}" /v ${slot} /f`);
        }
      } catch {
        // Slot doesn't exist, continue
      }
    }

    // Step 3: Register in registry for auto-load
    const registryResult = await registerAddinInRegistry(copyResult.path);

    // Step 4 (best-effort): If Excel is already running, load the add-in into the live instance
    const liveLoadResult = await tryLoadAddinIntoRunningExcel(copyResult.path);

    return {
      success: true,
      addinPath: copyResult.path,
      fileCopied: copyResult.copied,
      registrySlot: registryResult.slot,
      registryUpdated: registryResult.registered,
      liveLoadAttempted: liveLoadResult.attempted,
      liveLoadSucceeded: liveLoadResult.loaded,
      liveLoadReason: liveLoadResult.reason
    };

  } catch (error) {
    console.error('[Excel Add-in] Installation failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Uninstall the Excel Add-in:
 * 1. Remove from registry
 * 2. Delete from AddIns folder
 */
async function uninstallExcelAddin() {
  if (process.platform !== 'win32') {
    return { success: false, error: 'Windows only.' };
  }
  try {
    const baseKey = await getExcelOptionsRegistryKey();
    const addInsFolder = getAddInsFolder();
    const addinPath = path.join(addInsFolder, ADDIN_FILE_NAME);

    // Remove from registry (check all OPEN slots)
    const slots = ['OPEN', ...Array.from({ length: 10 }, (_, i) => `OPEN${i + 1}`)];

    for (const slot of slots) {
      try {
        const { stdout } = await execAsync(`reg query "${baseKey}" /v ${slot} 2>nul`);
        if (stdout.includes(ADDIN_FILE_NAME)) {
          await execAsync(`reg delete "${baseKey}" /v ${slot} /f`);
        }
      } catch {
        // Slot doesn't exist, continue
      }
    }

    // Delete file
    if (await fs.pathExists(addinPath)) {
      await fs.remove(addinPath);
    }

    return { success: true };
  } catch (error) {
    console.error('[Excel Add-in] Uninstall failed:', error.message);
    return { success: false, error: error.message };
  }
}

module.exports = { installExcelAddin, uninstallExcelAddin };
