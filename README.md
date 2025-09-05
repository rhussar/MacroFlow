# MacroFlow
AI powered VBA

# MacroFlow AI

🤖 Generate VBA macros for Excel using AI

## Quick Start

### 1. Backend Setup (Flask + OpenAI)
```bash
# Install Python dependencies
pip install -r requirements.txt

# Set your OpenAI API key (REQUIRED)
export OPENAI_API_KEY="your-actual-api-key-here"
# Or on Windows:
# set OPENAI_API_KEY=your-actual-api-key-here

# Start Flask backend
python app.py
```

### 2. Frontend Setup (Office Add-in)
```bash
# Install Node.js dependencies
npm install

# Start development server (Excel will open)
npm start
```

## Architecture

- **Backend**: Flask server with OpenAI integration and COM automation
- **Frontend**: Office.js Add-in for Excel integration
- **AI**: OpenAI GPT-4 for VBA code generation
- **Injection**: `win32com.client` for direct Excel VBA injection

## Development Commands

```bash
npm start          # Start Excel with add-in sideloaded
npm stop           # Stop development server
npm run validate   # Validate manifest.xml
npm run build      # Build for production
python app.py      # Start Flask backend
```

## Production Requirements

1. **Excel Trust Settings**: Enable "Trust access to VBA project object model"
2. **HTTPS Certificates**: Required for Office Add-in
3. **OpenAI API Key**: Set as environment variable

## API Endpoints

- `POST /generate` - Generate VBA from text prompt
- `POST /inject` - Inject VBA into active Excel workbook

## File Structure

```
MacroFlow-AI/
├── app.py                        # Flask backend
├── requirements.txt              # Python dependencies
├── AI-Frontend/
│   ├── manifest.xml             # Office Add-in config
│   ├── taskpane.html            # Main UI
│   ├── taskpane.js              # Frontend logic
│   └── style.css                # Styling
├── package.json                 # Node.js config
├── webpack.config.js            # Build configuration
└── README.md
```

## License

MIT
