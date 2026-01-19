const fs = require('fs-extra');
const path = require('path');
const { app } = require('electron');

/**
 * Get the path to the MacroFlowLoader.xlam file
 * Handles both development and production environments
 */
function getAddinSourcePath() {
  // In production, resources are in app.asar or app.asar.unpacked
  // In dev, they're in the project root
  const isDev = process.env.NODE_ENV === 'development';

  if (isDev) {
    // Development: resources folder is at project root
    return path.join(__dirname, '../resources/MacroFlowLoader.xlam');
  } else {
    // Production: Check if running from asar
    if (process.resourcesPath) {
      // First try app.asar.unpacked (if resources were marked to unpack)
      const unpackedPath = path.join(process.resourcesPath, 'app.asar.unpacked/resources/MacroFlowLoader.xlam');
      if (fs.existsSync(unpackedPath)) {
        return unpackedPath;
      }

      // Then try normal resources path
      const resourcesPath = path.join(process.resourcesPath, 'resources/MacroFlowLoader.xlam');
      if (fs.existsSync(resourcesPath)) {
        return resourcesPath;
      }
    }

    // Fallback to relative path
    return path.join(__dirname, '../resources/MacroFlowLoader.xlam');
  }
}

/**
 * Get the Excel XLSTART folder path
 * This is where Excel automatically loads add-ins from
 */
function getExcelXLSTARTPath() {
  const appData = process.env.APPDATA;
  if (!appData) {
    throw new Error('APPDATA environment variable not found');
  }

  return path.join(appData, 'Microsoft', 'Excel', 'XLSTART');
}

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
    // If either file doesn't exist or can't be read, they're not identical
    return false;
  }
}

/**
 * Install the Excel Add-in to the user's XLSTART folder
 * This makes it automatically load when Excel starts
 */
async function installExcelAddin() {
  try {
    // Get source and destination paths
    const sourcePath = getAddinSourcePath();
    const xlstartPath = getExcelXLSTARTPath();
    const destPath = path.join(xlstartPath, 'MacroFlowLoader.xlam');

    // Check if source file exists
    if (!await fs.pathExists(sourcePath)) {
      throw new Error(`MacroFlowLoader.xlam not found at ${sourcePath}`);
    }

    // Ensure XLSTART directory exists
    await fs.ensureDir(xlstartPath);

    // Check if destination file already exists and is identical
    if (await fs.pathExists(destPath)) {
      const identical = await areFilesIdentical(sourcePath, destPath);
      if (identical) {
        return {
          success: true,
          action: 'already_installed',
          path: destPath
        };
      }
    }

    // Copy the add-in file
    await fs.copy(sourcePath, destPath, { overwrite: true });

    return {
      success: true,
      action: 'installed',
      path: destPath
    };

  } catch (error) {
    console.error('[Excel Add-in] Installation failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

module.exports = { installExcelAddin };
