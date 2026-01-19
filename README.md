# MacroFlow

**A desktop sidecar application for Excel that enables VBA code injection via COM automation.**

MacroFlow runs as a standalone Electron app that docks to the right side of your screen, providing a Monaco code editor for writing and pushing VBA macros directly into Excel's VBA Editor.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                         YOUR SCREEN                              │
├─────────────────────────────────────┬───────────────────────────┤
│                                     │                           │
│           Microsoft Excel           │    MacroFlow Sidebar      │
│                                     │    (Electron + React)     │
│     ┌─────────────────────────┐     │                           │
│     │     Active Workbook     │     │    ┌─────────────────┐    │
│     │                         │◄────┼────│  Monaco Editor  │    │
│     │  ┌───────────────────┐  │ COM │    │                 │    │
│     │  │ MacroFlowModule   │  │     │    │  ' VBA Code     │    │
│     │  │                   │  │     │    │  Sub MyMacro()  │    │
│     │  │ Sub MyMacro()     │  │     │    │    ...          │    │
│     │  │   MsgBox "Hello"  │  │     │    │  End Sub        │    │
│     │  │ End Sub           │  │     │    │                 │    │
│     │  └───────────────────┘  │     │    └─────────────────┘    │
│     │                         │     │    [  Push to Excel  ]    │
│     └─────────────────────────┘     │                           │
│                                     │                           │
└─────────────────────────────────────┴───────────────────────────┘
```

**Tech Stack:**
- **Frontend:** React + Vite + Monaco Editor
- **Backend:** Electron (Node.js) with winax for COM automation
- **Build:** Electron Forge with Squirrel installer
- **Excel Integration:** COM via `winax.GetObject('Excel.Application')`

---

## Prerequisites

1. **Windows 10/11** (COM automation is Windows-only)
2. **Microsoft Excel** (Microsoft 365 or 2019+)
3. **Node.js 18+** with npm

---

## Quick Start

### Development Mode

```bash
cd App
npm install
npm run dev
```

This launches:
- Vite dev server on `http://localhost:5173`
- Electron app with hot reload
- DevTools for debugging

### Production Build

```bash
cd App
npm run make
```

Outputs installer to `App/out/make/squirrel.windows/x64/MacroFlow-x.x.x Setup.exe`

---

## Excel Configuration (Required)

**You must enable VBA project access or MacroFlow cannot inject code.**

1. Open Excel
2. Go to **File → Options → Trust Center → Trust Center Settings**
3. Select **Macro Settings**
4. Check: **Trust access to the VBA project object model**
5. Click OK and restart Excel

---

## How It Works

1. **Open Excel** with any workbook
2. **Launch MacroFlow** (via dev mode or installed app)
3. **Write VBA code** in the Monaco editor
4. **Click "Push to Excel"**
5. MacroFlow creates/updates `MacroFlowModule` in your workbook's VBA project

The connection uses `winax.GetObject('', 'Excel.Application')` to attach to the running Excel instance - no new Excel windows are created.

---

## Project Structure

```
MacroFlow/
├── App/                          # Main application
│   ├── electron/                 # Electron main process
│   │   ├── main.js              # App entry, window creation, single-instance lock
│   │   ├── preload.js           # Context bridge for IPC
│   │   ├── ipc-handlers.js      # All backend logic + Excel COM integration
│   │   └── excel-addin-installer.js  # Auto-installs .xlam to XLSTART
│   ├── src/                     # React frontend
│   │   ├── App.jsx              # Main component with Monaco editor
│   │   └── App.css              # Styles
│   ├── Resources/               # Excel add-in files
│   │   └── MacroFlowLoader.xlam # Ribbon button add-in
│   ├── package.json
│   └── forge.config.js          # Electron Forge build config
└── README.md
```

---

## Key Files

| File | Purpose |
|------|---------|
| `electron/main.js` | Single-instance lock, window positioning (safe zones), app lifecycle |
| `electron/ipc-handlers.js` | Excel COM connection via winax, VBA injection logic |
| `electron/preload.js` | Exposes `window.electronAPI` to renderer |
| `src/App.jsx` | React UI with Monaco editor and Push to Excel button |

---

## IPC API

```javascript
// Inject VBA code into Excel
window.electronAPI.injectCode(code)
// Returns: { success: boolean, message: string }

// Close the app
window.electronAPI.closeApp()
```

---

## Troubleshooting

### "Excel is not running"
- Open Excel with a workbook before clicking Push to Excel

### "Please open a workbook first"
- Create or open any Excel file (the workbook must be active)

### "Action Blocked: Trust access..."
- Enable VBA project access in Trust Center (see Excel Configuration above)

### Multiple windows opening
- Kill all `EXCEL.EXE` processes and restart
- The app uses `requestSingleInstanceLock()` to prevent duplicates

### Ghost Excel instances
- Run `taskkill /F /IM excel.exe` to clean up stale processes
- Restart Excel fresh

---

## Window Positioning

The sidebar uses "safe zone" gaps to avoid covering Excel controls:

```javascript
const TOP_GAP = 240;    // Leaves ribbon + column headers visible
const BOTTOM_GAP = 50;  // Leaves status bar + zoom visible
const SIDEBAR_WIDTH = 400;
```

Adjust these values in `electron/main.js` if needed.

---

## License

MIT
