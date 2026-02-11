const path = require('path');
const { FusesPlugin } = require('@electron-forge/plugin-fuses');
const { FuseV1Options, FuseVersion } = require('@electron/fuses');

module.exports = {
  packagerConfig: {
    icon: './assets/app-icon',
    asar: {
      unpackDir: 'node_modules',
    },
    // Copy each subfolder individually so they land directly under
    // process.resourcesPath (e.g. resourcesPath/models/, resourcesPath/llm/)
    extraResource: [
      './Resources/llm',
      './Resources/models',
      './Resources/MacroFlow.xlam',
    ],
    // Keep large build-time folders OUT of the asar
    ignore: [
      /^[/\\]Resources($|[/\\])/,
      /^[/\\]models($|[/\\])/,
      /^[/\\]scripts($|[/\\])/,
      // Strip GPU backends (CPU-only) and ARM64 — saves ~650 MB
      /^[/\\]node_modules[/\\]@node-llama-cpp[/\\]win-arm64($|[/\\])/,
      /^[/\\]node_modules[/\\]@node-llama-cpp[/\\]win-x64-cuda($|[/\\])/,
      /^[/\\]node_modules[/\\]@node-llama-cpp[/\\]win-x64-cuda-ext($|[/\\])/,
      /^[/\\]node_modules[/\\]@node-llama-cpp[/\\]win-x64-vulkan($|[/\\])/,
    ],
  },
  rebuildConfig: {},
  makers: [
    // 1. Wix (Creates a standard .msi installer that handles large files)
    {
      name: '@electron-forge/maker-wix',
      config: {
        language: 1033,
        manufacturer: 'MacroFlow',
        icon: path.join(__dirname, 'assets', 'app-icon.ico'),
        // Per-user install (AppData) — no UAC prompt, no admin required
        defaultInstallMode: 'perUser',
        // Unique UUID required by Windows for updates (do not change this once set)
        upgradeCode: '46200234-8025-4513-8877-111002345678',
        // Suppress per-user directory ICE checks (expected for LocalAppDataFolder)
        lightSwitches: ['-sice:ICE38', '-sice:ICE91', '-sice:ICE64'],
        // Redirect install dir from Program Files → %LOCALAPPDATA%\MacroFlow
        // (ProgramFilesFolder requires admin even with perUser scope)
        beforeCreate: (creator) => {
          creator.wixTemplate = creator.wixTemplate.replace(
            '{{ProgramFilesFolder}}',
            'LocalAppDataFolder',
          );
        },
      },
    },
    // 2. ZIP (Backup for testing)
    {
      name: '@electron-forge/maker-zip',
      platforms: ['win32'],
    },
  ],
  plugins: [
    {
      name: '@electron-forge/plugin-auto-unpack-natives',
      config: {},
    },
    // Fuses are used to enable/disable various Electron functionality
    // at package time, before code signing the application
    new FusesPlugin({
      version: FuseVersion.V1,
      [FuseV1Options.RunAsNode]: false,
      [FuseV1Options.EnableCookieEncryption]: true,
      [FuseV1Options.EnableNodeOptionsEnvironmentVariable]: false,
      [FuseV1Options.EnableNodeCliInspectArguments]: false,
      [FuseV1Options.EnableEmbeddedAsarIntegrityValidation]: true,
      [FuseV1Options.OnlyLoadAppFromAsar]: true,
    }),
  ],
};
