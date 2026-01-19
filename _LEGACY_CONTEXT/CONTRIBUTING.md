# Contributing to MacroFlow

Welcome to MacroFlow! This guide explains how the system works and how to contribute effectively.

## 📋 Table of Contents

- [How MacroFlow Works](#how-macroflow-works)
- [Architecture Overview](#architecture-overview)
- [Development Setup](#development-setup)
- [Project Structure](#project-structure)
- [Development Guidelines](#development-guidelines)
- [Testing](#testing)
- [Deployment](#deployment)

---

## How MacroFlow Works

MacroFlow is an AI-powered Office Add-in that generates VBA macros for Excel using a sophisticated dual-agent architecture.

### System Flow

```
User Input (Frontend)
    ↓
Intent Classification (AI)
    ↓
┌─────────────────┬─────────────────┐
│   Chat Agent    │   VBA Agent     │
│  (Conversation) │  (Code Gen)     │
└─────────────────┴─────────────────┘
    ↓                    ↓
Response            VBA Code
    ↓                    ↓
Frontend UI      Excel COM Injection
```

### Key Components

**1. Frontend (Office Add-in)**
- Task pane interface built with Office.js
- Monaco Editor for VBA code editing
- Real-time chat interface with AI
- Module management system
- Context extraction from Excel

**2. Backend (Flask API)**
- REST API endpoints for AI interactions
- Multi-agent AI architecture
- Excel COM automation (win32com)
- Sheet context analysis
- VBA module manipulation

**3. AI Agent System**
- **Intent Classifier**: Routes requests to appropriate agent
- **Chat Agent**: Handles conversations, questions, explanations
- **VBA Agent**: Generates and modifies VBA code with context awareness

---

## Architecture Overview

### Dual-Agent AI System

```python
# Request Flow
1. User sends prompt → Backend /generate endpoint
2. Intent Classifier analyzes prompt + context
3. Routes to:
   - Chat Agent (conversation/help) OR
   - VBA Agent (code generation/modification)
4. Agent processes with OpenAI GPT-4
5. Response returns to frontend
6. VBA code auto-injected into Excel (if applicable)
```

### Context-Aware Code Generation

The VBA Agent uses multiple context sources:
- **Cursor Position**: Detects which subroutine user is editing
- **Selection**: User-highlighted code for modification
- **Full Module**: Complete module code for analysis
- **Excel Sheet Data**: Current workbook/sheet context
- **Conversation History**: Previous messages for continuity

### Excel COM Integration

```python
# Direct VBA manipulation via win32com
pythoncom.CoInitialize()
excel = win32com.client.GetObject(None, "Excel.Application")
wb = excel.ActiveWorkbook
module = wb.VBProject.VBComponents.Add(1)  # Create module
module.CodeModule.AddFromString(vba_code)  # Inject code
```

### Frontend State Management

```javascript
// Main state object
state = {
    modules: [],           // All VBA modules
    activeModule: null,    // Currently selected module
    conversationHistory: [],
    editor: monacoInstance
}
```

---

## Development Setup

### Prerequisites

- **Python 3.8+** with pip
- **Node.js 16+** with npm
- **Microsoft Excel** (2016 or later)
- **OpenAI API key**

### Quick Start

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd MacroFlow
   ```

2. **Configure environment**
   ```bash
   # Create backend/.env file
   cp .env.example backend/.env
   # Add your OpenAI API key to backend/.env
   ```

3. **Install dependencies**
   ```bash
   # Backend
   cd backend
   pip install -r requirements.txt

   # Frontend
   cd ../frontend
   npm install
   ```

4. **Start the application**
   ```bash
   # Windows (from project root)
   ./start.bat

   # Linux/macOS
   scripts/start.sh

   # Or start backend only
   cd backend && py app.py
   ```

5. **Excel configuration**
   - File → Options → Trust Center → Trust Center Settings
   - Macro Settings → Check "Trust access to VBA project object model"
   - Install dev certificates: `cd frontend && npx office-addin-dev-certs install`

---

## Project Structure

```
MacroFlow/
├── backend/                    # Python Flask backend
│   ├── agents/                # AI agent system
│   │   ├── __init__.py
│   │   ├── base_agent.py      # Base class for all agents
│   │   ├── chat_agent.py      # Conversational AI agent
│   │   ├── intent_classifier.py  # Intent routing logic
│   │   └── vba_agent.py       # VBA code generation agent
│   ├── app.py                 # Main Flask application
│   ├── config.py              # Environment-based configuration
│   ├── requirements.txt       # Python dependencies
│   ├── .env                   # Environment variables (local only, gitignored)
│   └── start-backend.bat      # Backend-only startup script
│
├── frontend/                   # Office Add-in
│   ├── src/
│   │   ├── taskpane.html      # Main UI layout
│   │   ├── taskpane.js        # Frontend logic (2,240 lines)
│   │   └── style.css          # Styling
│   ├── assets/
│   │   └── icons/             # Office ribbon icons (16, 32, 64, 80px)
│   ├── manifest.xml           # Office Add-in configuration
│   ├── package.json           # Node.js dependencies
│   ├── server.js              # HTTPS development server
│   └── webpack.config.js      # Build configuration
│
├── docs/                      # Documentation
│   ├── API.md                 # API endpoint reference
│   └── SETUP.md               # Detailed setup guide
│
├── scripts/                   # Automation scripts
│   ├── start.sh               # Linux/macOS startup
│   └── deploy.bat             # Production build script
│
├── tests/                     # Test files
│   ├── test_backend.py        # Backend unit tests
│   └── test_frontend.js       # Frontend tests
│
├── start.bat                  # Windows startup (root)
├── .env.example               # Environment template
├── .gitignore                 # Git ignore patterns
├── LICENSE                    # MIT License
├── CHANGELOG.md               # Version history
├── CONTRIBUTING.md            # This file
└── README.md                  # Project overview
```

---

## Development Guidelines

### Code Style

**Python (Backend)**
- Follow PEP 8
- Use type hints: `def process(self, prompt: str) -> Dict[str, Any]:`
- Add docstrings to all functions and classes
- Max line length: 100 characters

**JavaScript (Frontend)**
- Use ES6+ features (async/await, arrow functions)
- camelCase for variables and functions
- Add JSDoc comments for complex functions
- Max line length: 100 characters

### Commit Conventions

Use descriptive commit messages:
- `Add: New VBA code optimization feature`
- `Fix: Excel COM connection timeout issue`
- `Update: Chat agent response formatting`
- `Remove: Deprecated configuration files`
- `Docs: Update API endpoint documentation`

### Agent Development

**Creating a New Agent:**

1. Inherit from `BaseAgent` class
2. Implement the `process()` method
3. Add to agent initialization in `app.py`

```python
from agents.base_agent import BaseAgent

class MyNewAgent(BaseAgent):
    def process(self, prompt: str, context: Optional[Dict] = None,
                conversation_history: Optional[List] = None) -> Dict[str, Any]:
        # Your agent logic here
        response = self._call_openai(messages, temperature=0)
        return {
            'type': 'custom',
            'content': response
        }
```

### Frontend Development

**Office Add-in Guidelines:**
- All API calls use `http://localhost:5000` (backend)
- Monaco Editor instance stored in `state.editor`
- Chat messages stored in `state.conversationHistory`
- Module data cached in `state.modules`

**Key Functions:**
- `generateMacro()` - Main AI interaction handler
- `insertVBACode()` - Smart code insertion logic
- `loadModules()` - Fetch VBA modules from Excel
- `getCurrentSubroutine()` - Detect cursor position context

---

## Testing

### Backend Tests

```bash
cd backend
python -m pytest tests/
```

**Test Coverage:**
- Agent system unit tests
- API endpoint integration tests
- Excel COM mock tests
- Error handling scenarios

### Frontend Tests

```bash
cd frontend
npm test
npm run validate    # Validate Office manifest
npm run lint       # ESLint checks
```

### Manual Testing Checklist

- [ ] Start application (both servers running)
- [ ] Open Excel with test workbook
- [ ] Generate VBA code via chat
- [ ] Verify code injection into Excel
- [ ] Test module creation/deletion
- [ ] Test subroutine execution
- [ ] Test error scenarios (no Excel, VBA access disabled)

---

## Deployment

### Cloud Deployment Considerations

**Backend (Flask)**
- Deploy to: AWS, Azure, Heroku, DigitalOcean
- Set environment variables on hosting platform
- Use production WSGI server (Gunicorn, uWSGI)
- Enable CORS for frontend domain
- Set `FLASK_ENV=production`

**Frontend (Static Files)**
- Build: `cd frontend && npm run build`
- Deploy `dist/` folder to static hosting
- Requires HTTPS (Office Add-in requirement)
- Update `manifest.xml` URLs to production domain

**Office Add-in Distribution**
- Upload manifest to Microsoft AppSource
- Or use organizational catalog for internal deployment
- Users install via Office Add-ins store

### Environment Variables

**Production `.env`:**
```bash
OPENAI_API_KEY=your_production_key
FLASK_ENV=production
BACKEND_HOST=0.0.0.0
BACKEND_PORT=5000
```

### Security Checklist

- [ ] Environment variables not committed to git
- [ ] HTTPS enabled on all endpoints
- [ ] CORS restricted to production domain
- [ ] API rate limiting configured
- [ ] Error messages don't expose sensitive data
- [ ] Excel VBA project access documented

---

## Technology Stack

### Backend
- **Flask 3.0.0** - Web framework
- **Flask-CORS 4.0.0** - Cross-origin requests
- **OpenAI SDK 1.30.0+** - GPT-4 API integration
- **pywin32 311** - Windows COM automation
- **python-dotenv 1.0.0** - Environment variable management

### Frontend
- **Office.js** - Office Add-in platform
- **Monaco Editor** - Code editor with VBA syntax highlighting
- **Express.js** - Development HTTPS server
- **Webpack** - Build system

### Development Tools
- **npm** - Package management
- **pip** - Python package management
- **git** - Version control

---

## API Architecture

### Core Endpoints

| Endpoint | Method | Purpose | Request | Response |
|----------|--------|---------|---------|----------|
| `/generate` | POST | Generate/modify VBA | `{prompt, context, conversation_history}` | `{type, vba_code, explanation}` |
| `/inject` | POST | Inject VBA into Excel | `{macro}` | `{success, message}` |
| `/list-modules` | GET | List all VBA modules | None | `{workbook, modules[]}` |
| `/create-module` | POST | Create new module | `{name, content}` | `{success, message}` |
| `/update-module` | POST | Update module content | `{name, content}` | `{success, message}` |
| `/delete-module` | POST | Delete module | `{name}` | `{success, message}` |
| `/run-subroutine` | POST | Execute VBA sub | `{module, subroutine}` | `{success, message}` |
| `/get-sheet-context` | GET | Get Excel sheet data | None | `{sheet_name, data_sample, ...}` |

### Request Flow Example

```javascript
// Frontend request
const response = await fetch('http://localhost:5000/generate', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
        prompt: "Create a macro to highlight cells > 100",
        context: {
            activeModule: "Module1",
            fullCode: "...",
            cursorPosition: {...},
            sheetContext: {...}
        },
        conversation_history: [...]
    })
});

// Backend processing
// 1. Intent Classifier determines intent
// 2. Routes to VBA Agent
// 3. VBA Agent generates code with context
// 4. Returns structured response

// Frontend receives
{
    type: 'vba',
    has_vba: true,
    vba_code: 'Sub HighlightCells()...',
    explanation: 'This macro highlights...',
    classification: {
        intent: 'vba_generation',
        confidence: 'high'
    }
}
```

---

## Performance Optimization

### Backend
- Sheet context cached for 30 seconds
- COM initialization per thread
- Efficient VBA code parsing
- OpenAI API call optimization

### Frontend
- Monaco Editor lazy loading
- Debounced auto-save (2 seconds)
- Module list caching
- Efficient DOM updates

---

## Troubleshooting

### Common Issues

**Backend won't start:**
- Check Python version: `py --version` (need 3.8+)
- Verify OpenAI API key in `backend/.env`
- Check port 5000 availability

**Frontend won't load:**
- Install dev certificates: `npx office-addin-dev-certs install`
- Check port 3000 availability
- Verify HTTPS is working

**VBA injection fails:**
- Enable "Trust access to VBA project object model" in Excel
- Ensure Excel is running
- Check if workbook is open

**Agent responses empty:**
- Verify OpenAI API key is valid
- Check API rate limits
- Review backend console for errors

---

## License

By contributing to MacroFlow, you agree that your contributions will be licensed under the MIT License.

---

**Questions?** Review the docs or check existing code patterns. The codebase is well-documented and follows consistent patterns throughout.
