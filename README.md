# MacroFlow

**AI-Powered VBA Macro Generation for Excel**

MacroFlow is an intelligent Excel Add-In that generates, edits, and manages VBA macros using AI.
This README focuses on developer setup for running MacroFlow locally.

---

## 🚀 Developer Setup Guide (Windows)

Follow these steps to fully install and run MacroFlow on a new machine.

---

### 1. Install Microsoft Excel

- Microsoft 365 or Excel 2019+
- Sign in with your Syracuse email


### 2. Install Core Tooling

#### 2.1 Python

Download from: https://python.org

✔ Check "Add Python to PATH"

Verify:

```bash
python --version
pip --version
```

#### 2.2 Node.js + npm

Download LTS from: https://nodejs.org

Fix script execution policy:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Verify:

```bash
node -v
npm -v
```

---

### 3. Backend Setup (Flask API)

```bash
cd backend
pip install -r requirements.txt
```

**Environment Variables**

1. Rename `backend/.env.example` → `.env`
2. Add your OpenAI key from OpenAI API platform:

```
OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxx
```

Ensure `.env` stays inside the `backend` folder.

---

### 4. Frontend / Add-In Setup

Cd into your MacroFlow folder

```bash
npm install
```

**Install Office HTTPS certificates**

If you ever see an "add-in error" in the Excel taskpane, you need to complete this step:

Run PowerShell as Administrator:

```bash
npx office-addin-dev-certs install
```

Accept all prompts.

**Start Dev Environment**

Run:

```bash
Start.bat
```

Backend should show:

```
Serving Flask app 'app'
Running on http://localhost:5000
```

---

### 5. Enable macros in Excel

If the "create module" button does not work, follow these steps:

Enable access to VBA automation:

**File → Options → Trust Center → Trust Center Settings → Macro Settings**

✔ Check "Trust access to the VBA project object model"

---

### 6. Running MacroFlow

1. Run `Start.bat`
2. Open Excel

---

### 7. Troubleshooting

**Add-In won't load**

- Ensure you installed certificates as Administrator
- Make sure backend is on `http://localhost:5000`

**Buttons do nothing**

- Check that:
  - Node dev server is running
  - Certificates are installed
  - Excel VBA project access is enabled

**Backend errors**

```bash
pip install -r backend/requirements.txt --force-reinstall
```

---

## Project Structure (Minimal Overview)

```
MacroFlow/
├── backend/        # Flask API & agents
├── frontend/       # Office Add-in (taskpane UI)
├── manifest.xml    # Excel Add-in manifest
├── Start.bat       # Dev startup script
└── README.md
```

---

## Environment Variables

Create a `.env` inside `/backend/`:

```
OPENAI_API_KEY=your_key_here
```

---

## 📄 License

MIT License. See LICENSE.

---

## 📬 Support

For issues, open a GitHub issue or contact the MacroFlow team.

---
