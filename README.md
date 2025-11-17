# MacroFlow

**AI-Powered VBA Macro Generation for Excel**

MacroFlow is an intelligent Office Add-in that enables users to generate, edit, and manage VBA macros using artificial intelligence. Built with a modern dual-architecture system featuring specialized AI agents for different tasks.

![MacroFlow Interface](docs/images/interface-preview.png)

## 🚀 Quick Start

### Prerequisites
- Python 3.8+ with pip
- Node.js 16+ with npm
- Microsoft Excel (2016 or later)
- OpenAI API key

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/macroflow-ai/macroflow-ai.git
   cd macroflow-ai
   ```

2. **Environment setup**
   ```bash
   cp .env.example .env
   # Edit .env with your OpenAI API key
   ```

3. **Install dependencies**
   ```bash
   # Backend dependencies
   cd backend
   pip install -r requirements.txt
   
   # Frontend dependencies  
   cd ../frontend
   npm install
   ```

4. **Start the application**
   ```bash
   # From project root
   scripts/start.bat    # Windows
   scripts/start.sh     # Linux/macOS
   ```

## 🏗️ Architecture

### Backend (Python Flask)
- **Multi-Agent System**: Intelligent task routing with specialized agents
- **Intent Classifier**: Determines user intent and routes requests appropriately
- **Chat Agent**: Handles conversational interactions and explanations
- **VBA Agent**: Generates and modifies VBA code with context awareness
- **Excel Integration**: Direct COM automation for seamless VBA injection

### Frontend (Office Add-in)
- **VSCode-inspired Interface**: Professional editor with syntax highlighting
- **Real-time Chat**: Streaming conversations with AI agents
- **Module Management**: Create, edit, and organize VBA modules
- **Context Awareness**: Integrates with active Excel workbook data

## ✨ Key Features

### 🤖 AI-Powered Code Generation
- Context-aware VBA macro generation
- Intelligent code modification and enhancement
- Natural language to VBA code translation
- Function-specific targeting for existing code

### 💬 Intelligent Chat System
- Multi-agent conversation routing
- Conversation history tracking
- Code explanation and documentation
- Interactive help and guidance

### 🔧 Advanced Excel Integration
- Direct VBA module injection
- Real-time workbook context reading
- Sheet data analysis and integration
- Macro execution from ribbon interface

### 🎨 Professional UI/UX
- Monaco Editor integration with VBA syntax highlighting
- Collapsible chat panel with streaming responses
- Professional ribbon integration
- Context-sensitive help and tooltips

## 📁 Project Structure

```
MacroFlow/
├── backend/              # Python Flask backend
│   ├── agents/          # AI agent system
│   │   ├── __init__.py
│   │   ├── base_agent.py
│   │   ├── chat_agent.py
│   │   ├── intent_classifier.py
│   │   └── vba_agent.py
│   ├── app.py           # Main Flask application
│   ├── config.py        # Configuration management
│   └── requirements.txt # Python dependencies
├── frontend/            # Office Add-in
│   ├── src/            # Source files
│   │   ├── taskpane.html
│   │   ├── taskpane.js
│   │   └── style.css
│   ├── assets/         # Static assets
│   │   └── icons/      # PNG icons for Office
│   ├── manifest.xml    # Office Add-in manifest
│   ├── package.json    # Node.js dependencies
│   ├── server.js       # Development server
│   └── webpack.config.js # Build configuration
├── docs/               # Documentation
│   ├── README.md       # Project overview
│   ├── SETUP.md        # Detailed setup guide
│   └── API.md          # API documentation
├── scripts/            # Build and deployment
│   ├── start.bat       # Windows startup script
│   ├── start.sh        # Linux/macOS startup script
│   └── deploy.bat      # Production deployment
├── config/             # Environment configurations
│   ├── development.json
│   └── production.json
├── tests/              # Test files
│   ├── test_backend.py
│   └── test_frontend.js
├── .env.example        # Environment template
└── .gitignore          # Git ignore rules
```

## 🛠️ Development

### Running in Development Mode

**Backend:**
```bash
cd backend
python app.py
# Server starts at http://localhost:5000
```

**Frontend:**
```bash
cd frontend
node server.js          # Development server
npm run dev             # Webpack dev server with hot reload
```

**Office Add-in:**
```bash
cd frontend
npm start               # Start Office add-in debugging
```

### Building for Production

```bash
cd frontend
npm run build           # Build frontend assets

# Or use deployment script
scripts/deploy.bat      # Windows deployment
```

### Testing

```bash
# Backend tests
python tests/test_backend.py

# Frontend validation
cd frontend
npm run validate        # Validate Office manifest
npm run lint           # ESLint checks
```

## 🔧 Configuration

### Environment Variables
**Required:** Set your OpenAI API key as an environment variable:
```bash
# Windows
set OPENAI_API_KEY=your_openai_api_key_here

# macOS/Linux
export OPENAI_API_KEY=your_openai_api_key_here

# Or create .env file in backend/ directory
OPENAI_API_KEY=your_openai_api_key_here
FLASK_ENV=development

# Server Configuration
BACKEND_HOST=localhost
BACKEND_PORT=5000
FRONTEND_HOST=localhost
FRONTEND_PORT=3000

# Office Add-in
OFFICE_ADDIN_ID=fba090d8-b4fc-4f86-a421-fa32126eda1a
OFFICE_ADDIN_VERSION=1.0.0.0
```

### Excel Configuration
1. **Enable VBA Project Access:**
   - File → Options → Trust Center → Trust Center Settings
   - Macro Settings → Check "Trust access to VBA project object model"
   - Restart Excel

2. **Install Development Certificates:**
   ```bash
   cd frontend
   npx office-addin-dev-certs install
   ```

## 📚 API Reference

### Core Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/generate` | POST | Generate VBA code using AI agents |
| `/inject` | POST | Inject VBA code into Excel workbook |
| `/list-modules` | GET | List all VBA modules in workbook |
| `/create-module` | POST | Create new VBA module |
| `/update-module` | POST | Update existing VBA module content |
| `/delete-module` | POST | Delete VBA module |
| `/run-subroutine` | POST | Execute VBA subroutine |
| `/get-sheet-context` | GET | Get current Excel sheet context |
| `/get-active-module` | GET | Get active module for ribbon |

For detailed API documentation, see [docs/API.md](docs/API.md).

## 🚀 Deployment

### Production Deployment

1. **Configure environment:**
   ```bash
   cp .env.example .env
   # Set production values in .env
   ```

2. **Build assets:**
   ```bash
   cd frontend
   npm run build
   ```

3. **Deploy backend:**
   - Deploy `backend/` folder to your Flask hosting service
   - Ensure Python dependencies are installed

4. **Deploy frontend:**
   - Deploy built assets to HTTPS web server
   - Update manifest.xml URLs to production domain

5. **Office Add-in distribution:**
   - Upload to Microsoft AppSource
   - Or distribute via organizational catalog

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

See [CONTRIBUTING.md](CONTRIBUTING.md) for detailed contribution guidelines.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support & Troubleshooting

### Common Issues

1. **Add-in not loading:** Check Office add-in certificates and trust settings
2. **VBA access denied:** Enable "Trust access to VBA project object model"
3. **Backend connection issues:** Verify Flask server is running on port 5000
4. **HTTPS certificate issues:** Re-install Office development certificates

### Getting Help

- 📖 Read the [Setup Guide](docs/SETUP.md)
- 🐛 Report bugs via [GitHub Issues](https://github.com/macroflow-ai/macroflow-ai/issues)
- 💬 Join our [Discord Community](https://discord.gg/macroflow)
- 📧 Contact: support@macroflow.ai

## 🙏 Acknowledgments

- OpenAI for GPT-4 API capabilities
- Microsoft Office Add-in development platform
- Monaco Editor for code editing interface
- Flask and Express.js frameworks

---

**Made with ❤️ by the MacroFlow Team**