const { FusesPlugin } = require('@electron-forge/plugin-fuses');
const { FuseV1Options, FuseVersion } = require('@electron/fuses');

module.exports = {
  packagerConfig: {
    icon: './assets/app-icon',
    prune: true,
    asar: {
      unpack: '*.node'
    },
    // Keep the packaged app lean: runtime gets helper/resources from extraResource.
    ignore: [
      /^\/out($|\/)/,
      /^\/dist-electron($|\/)/,
      /^\/window-monitor\.log$/,
      /^\/Resources($|\/)/,
      /^\/native\/window-focus-helper($|\/)/,
      /^\/node_modules\/@electron-forge($|\/)/,
      /^\/node_modules\/@electron\/fuses($|\/)/,
      /^\/node_modules\/electron-rebuild($|\/)/,
      /^\/node_modules\/(vite|@vitejs|tailwindcss|@tailwindcss|concurrently|wait-on|cross-env|tw-animate-css)($|\/)/
    ],
    extraResource: [
      './Resources',
      './native/window-focus-helper/bin-helper'
    ]
  },
  rebuildConfig: {},
  makers: [
    {
      name: '@electron-forge/maker-squirrel',
      // Keep this in sync with `app.setAppUserModelId(...)` in `electron/main.js` so
      // Windows taskbar grouping/pinning uses the right icon.
      config: { setupIcon: './assets/app-icon.ico', appId: 'com.macroflow.desktop' },
    },
    {
      name: '@electron-forge/maker-zip',
      platforms: ['darwin'],
    },
    {
      name: '@electron-forge/maker-deb',
      config: {},
    },
    {
      name: '@electron-forge/maker-rpm',
      config: {},
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


