# MacroFlow

MacroFlow is a Windows Excel sidecar that generates, runs, and organizes VBA macros. Create uses Claude Sonnet 5 through a Cloudflare Worker that holds the Anthropic API key.

## Download (Windows)

Installer:

[https://pub-a7aa338dce944ce383fc182f58a87366.r2.dev/installer/MacroFlow-Setup.exe](https://pub-a7aa338dce944ce383fc182f58a87366.r2.dev/installer/MacroFlow-Setup.exe)

Paste that URL on macroflow.ai after you publish a build.

1. Run `MacroFlow-Setup.exe`.
2. The installer copies `MacroFlow.xlam` into Excel’s XLSTART folder and tries to load it into a running Excel instance.
3. Open Excel (or restart it if the ribbon is missing) and click **Home → MacroFlow**.

Excel requirements:

- Excel must be running for MacroFlow to talk to workbooks.
- Enable **Trust access to the VBA project object model**: File → Options → Trust Center → Trust Center Settings → Macro Settings.

## Publish a new installer

### Via GitHub Actions (recommended)

1. Add these repository secrets: `AWS_ACCESS_KEY_ID`, `AWS_SECRET_ACCESS_KEY`, `S3_ENDPOINT`, `S3_BUCKET`.
2. Push a tag (`v*`) or run **Build & Sign Windows** via workflow_dispatch.
3. CI builds on `windows-2022`, uploads the `windows-installer` artifact, and (when secrets are set) publishes to R2 as `installer/MacroFlow-Setup.exe`.

Public URL after publish:

`https://pub-a7aa338dce944ce383fc182f58a87366.r2.dev/installer/MacroFlow-Setup.exe`

### Locally

On a Windows machine with signing credentials (optional) and R2 keys:

```bash
cd App
npm install
npm run make
npm run upload:update
```

`App/scripts/publish-update.js` uploads:

- `installer/MacroFlow-Setup.exe` — stable public download URL
- `updates/RELEASES` and the `.nupkg` — Squirrel auto-update feed

Required env vars: `AWS_ACCESS_KEY_ID`, `AWS_SECRET_ACCESS_KEY`, `S3_ENDPOINT`, `S3_BUCKET`.

## AI proxy (Claude Sonnet 5)

The desktop app never contains an Anthropic key. Deploy the Worker first:

```bash
cd worker
npx wrangler login
npx wrangler secret put ANTHROPIC_API_KEY
npx wrangler deploy
```

Default app URL: `https://macroflow-ai.macroflowai.workers.dev`

Override with `MACROFLOW_AI_URL` if your Worker hostname differs. Details: [worker/README.md](worker/README.md).

## Development Mode

```bash
cd App
npm install
npm run dev
```

This starts Vite on `http://localhost:5173` and launches the Electron app.

## Files

- `docs/excel-bridge-api.md` documents the current `window.excel` API exposed to the UI.
- `App/package.json` defines the app metadata, scripts, and dependencies.
- `App/Resources/MacroFlow.xlam` is the Excel add-in copied into XLSTART.
- `App/electron/main.js` creates the Electron window and manages app lifecycle.
- `App/electron/excel-addin-installer.js` installs and uninstalls the Excel add-in.
- `App/electron/llm-client.js` calls the Cloudflare Worker for VBA generation.
- `worker/` is the Anthropic proxy (Claude Sonnet 5).
