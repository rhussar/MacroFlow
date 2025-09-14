# MacroFlow AI - Setup Guide

## Prerequisites

### Required Software
- **Python 3.8+** with pip
- **Node.js 16+** with npm
- **Microsoft Excel** (2016 or later)
- **Git** for version control

### Office Add-in Development
- Office Add-in development certificates installed
- Excel Trust Center configured for VBA access

## Development Setup

### 1. Clone Repository
```bash
git clone https://github.com/macroflow-ai/macroflow-ai.git
cd macroflow-ai
```

### 2. Environment Configuration
```bash
# Copy environment template
cp .env.example .env

# Edit .env file with your OpenAI API key
# OPENAI_API_KEY=your_openai_api_key_here
```

### 3. Backend Setup
```bash
cd backend
pip install -r requirements.txt
python app.py
```
Backend will start at `http://localhost:5000`

### 4. Frontend Setup
```bash
cd frontend
npm install
```

**Development Mode (with debugging):**
```bash
npm start
# Answer "Y" when prompted for localhost loopback
# Excel will launch automatically with MacroFlow loaded
```

**Production Mode (without debug dialogs):**
```bash
npm run sideload
# Restart Excel manually
# Add-in will be available without debug prompts
```

### 5. Excel Trust Center Configuration

**Enable VBA Project Access:**
1. Excel → File → Options → Trust Center → Trust Center Settings
2. Macro Settings → Check "Trust access to VBA project object model"
3. Restart Excel

## Office Add-in Certificates

### Install Development Certificates
```bash
cd frontend
npx office-addin-dev-certs install
```

### Verify Certificates
```bash
npx office-addin-dev-certs verify
```

## Project Structure

```
MacroFlow/
├── backend/           # Python Flask backend
│   ├── agents/        # AI agent system  
│   ├── api/          # API routes
│   ├── utils/        # Utility functions
│   ├── app.py        # Main Flask app
│   └── config.py     # Configuration
├── frontend/         # Office Add-in
│   ├── src/          # Source files
│   ├── assets/       # Static assets
│   └── dist/         # Built files
├── docs/             # Documentation
├── scripts/          # Build scripts
└── config/           # Environment configs
```

## Development Workflow

### Starting Development
```bash
# Terminal 1: Backend
cd backend && python app.py

# Terminal 2: Frontend  
cd frontend && npm start
```

### Code Changes
- **Backend**: Flask auto-reloads on Python file changes
- **Frontend**: Use `npm run dev` for webpack dev server with hot reload

### Testing
```bash
# Validate Office manifest
npm run validate

# Run tests
python tests/test_runner.py
```

## Troubleshooting

### Common Issues

**Add-in Registration Error:**
- Clear Office cache: Delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
- Restart Excel completely
- Re-run `npm start`

**VBA Access Denied:**
- Check Excel Trust Center settings
- Ensure "Trust access to VBA project object model" is enabled

**HTTPS Certificate Issues:**
- Re-install certificates: `npx office-addin-dev-certs install --machine`
- Clear browser cache
- Restart Excel

**Debug Dialog Issues:**
If you see "To debug your Event-based handler" dialogs:
- Use production sideload: `npm run sideload` instead of `npm start`
- Remove debug registrations: `npm run remove` then `npm run sideload`
- Restart Excel after changing sideload method

**Backend Connection Issues:**
- Verify Python Flask server is running on port 5000
- Check firewall settings
- Ensure CORS is properly configured

### Debug Mode
```bash
# Backend debug mode
FLASK_ENV=development python backend/app.py

# Frontend debug mode
npm run dev
```

## Production Deployment

See [deploy.bat](../scripts/deploy.bat) for production deployment instructions.

## Support

For issues and questions:
1. Check [troubleshooting](#troubleshooting) section
2. Search existing [GitHub issues](https://github.com/macroflow-ai/macroflow-ai/issues)
3. Create new issue with detailed description