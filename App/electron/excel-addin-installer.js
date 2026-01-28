const fs = require('fs-extra');
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');

const execAsync = promisify(exec);

// ============================================================================
// PATH HELPERS
// ============================================================================

/**
 * Get the path to the MacroFlowLoader.xlam file
 * Handles both development and production environments
 */
function getAddinSourcePath() {
  const isDev = process.env.NODE_ENV === 'development';

  if (isDev) {
    // Development: resources folder is at project root
    return path.join(__dirname, '../Resources/MacroFlowLoader.xlam');
  } else {
    // Production: extraResource copies to process.resourcesPath/resources
    if (process.resourcesPath) {
      const candidates = [
        path.join(process.resourcesPath, 'Resources', 'MacroFlowLoader.xlam'),
        path.join(process.resourcesPath, 'resources', 'MacroFlowLoader.xlam'),
      ];
      for (const candidate of candidates) {
        if (fs.existsSync(candidate)) {
          return candidate;
        }
      }
    }
    // Fallback
    return path.join(__dirname, '../Resources/MacroFlowLoader.xlam');
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
  const destPath = path.join(addInsFolder, 'MacroFlowLoader.xlam');

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
  const baseKey = 'HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Options';

  // Check OPEN first
  try {
    const { stdout } = await execAsync(`reg query "${baseKey}" /v OPEN 2>nul`);
    // OPEN exists, check if it's our add-in
    if (stdout.includes('MacroFlowLoader.xlam')) {
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
      if (stdout.includes('MacroFlowLoader.xlam')) {
        return { slot: valueName, alreadyRegistered: true };
      }
    } catch {
      // This slot is free
      return { slot: valueName, alreadyRegistered: false };
    }
  }

  // All slots taken, overwrite OPEN as fallback
  return { slot: 'OPEN', alreadyRegistered: false };
}

/**
 * Register the add-in in the Windows registry for Excel auto-load
 * @param {string} destinationPath - The path where the add-in was copied to (in AddIns folder)
 */
async function registerAddinInRegistry(destinationPath) {
  const baseKey = 'HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Options';

  // Find available slot
  const { slot, alreadyRegistered } = await findNextOpenSlot();

  if (alreadyRegistered) {
    return { registered: false, slot, reason: 'already_registered' };
  }

  // Format: /R "C:\Users\...\AppData\Roaming\Microsoft\AddIns\MacroFlowLoader.xlam"
  // The /R switch tells Excel to load as a hidden add-in (not a visible workbook)
  // We must escape the inner quotes for the Windows reg command
  const valueData = `/R \\"${destinationPath}\\"`;

  // Build the reg command
  // REG_SZ is the string type for registry values
  const regCommand = `reg add "${baseKey}" /v ${slot} /t REG_SZ /d "${valueData}" /f`;

  try {
    await execAsync(regCommand);
    return { registered: true, slot, path: destinationPath };
  } catch (error) {
    throw new Error(`Registry write failed: ${error.message}`);
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
  try {
    // Step 1: Get source path
    const sourcePath = getAddinSourcePath();

    if (!await fs.pathExists(sourcePath)) {
      throw new Error(`MacroFlowLoader.xlam not found at ${sourcePath}`);
    }

    // Step 2: Copy to AddIns folder
    const copyResult = await copyAddinToAddInsFolder(sourcePath);

    // Step 3: Register in registry for auto-load
    const registryResult = await registerAddinInRegistry(copyResult.path);

    return {
      success: true,
      addinPath: copyResult.path,
      fileCopied: copyResult.copied,
      registrySlot: registryResult.slot,
      registryUpdated: registryResult.registered
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
  try {
    const baseKey = 'HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Options';
    const addInsFolder = getAddInsFolder();
    const addinPath = path.join(addInsFolder, 'MacroFlowLoader.xlam');

    // Remove from registry (check all OPEN slots)
    const slots = ['OPEN', ...Array.from({ length: 10 }, (_, i) => `OPEN${i + 1}`)];

    for (const slot of slots) {
      try {
        const { stdout } = await execAsync(`reg query "${baseKey}" /v ${slot} 2>nul`);
        if (stdout.includes('MacroFlowLoader.xlam')) {
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
