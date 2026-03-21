# MacroFlow

MacroFlow generates VBA with a local on-device model. On first use of Create, the app will prompt the user to install the local AI runtime and download the configured model. App updates and licensing remain separate network-backed features.

## Install (Windows)

1) Build the installer:

```bash
cd App
npm install
npm run make
```

2) Run the installer:

- Open `App/out/make/squirrel.windows/x64/`
- Double-click `MacroFlow-x.x.x Setup.exe`

3) First use of Create:

- Open Create inside MacroFlow
- If local AI is not installed yet, click `Install Local AI`
- MacroFlow will download the local AI runtime and the configured local code model

## Development Mode

```bash
cd App
npm install
npm run dev
```

This starts Vite on `http://localhost:5173` and launches the Electron app.

## Local AI

- Provider: `Ollama`
- Default model: `qwen2.5-coder:3b`
- Runtime status and setup are available in MacroFlow Settings
- VBA generation stays on the local machine and talks only to a loopback Ollama API
- Ollama cloud features are disabled for MacroFlow's managed runtime

## Files 


- `docs/excel-bridge-api.md` documents the current `window.excel` API exposed to the UI.
- `App/package.json` defines the app metadata, scripts, and dependencies.
- `App/package-lock.json` locks dependency versions for deterministic installs.

- `App/forge.config.js` configures Electron Forge packaging and makers.
- `App/vite.config.js` configures the Vite build for the renderer.
- `App/index.html` is the renderer HTML entry point.
- `App/components.json` holds UI tooling configuration.
- `App/test-connection.js` is a standalone winax/Excel COM smoke test (requires matching Node version; otherwise test from Electron).
- `App/Resources/MacroFlowLoader.xlam` is the Excel add‑in loaded by the installer.

- `App/electron/main.js` creates the Electron window and manages app lifecycle.
- `App/electron/preload.js` exposes a safe IPC bridge to the renderer.
- `App/electron/ipc-handlers.js` routes renderer IPC calls to backend logic.
- `App/electron/excel-bridge.js` implements Excel COM automation and VBA helpers.
- `App/electron/excel-addin-installer.js` installs and registers the Excel add‑in.

- `App/src/main.jsx` boots the React app in the renderer.
- `App/src/index.css` defines global styles and Tailwind base styles.
- `App/src/App.jsx` is the current React UI entry component.
