# MacroFlow AI - Documentation

## Overview

MacroFlow AI is an intelligent Office Add-in that generates VBA macros for Excel using artificial intelligence. The application features a dual-architecture system with specialized AI agents for different tasks.

## Architecture

### Backend (Python Flask)
- **AI Agent System**: Multi-agent architecture for intelligent task routing
- **Intent Classifier**: Determines user intent and routes requests appropriately  
- **Chat Agent**: Handles conversational interactions and explanations
- **VBA Agent**: Generates and modifies VBA code with context awareness
- **Excel Integration**: Direct COM automation for seamless VBA injection

### Frontend (Office Add-in)
- **Task Pane Interface**: VSCode-inspired editor with syntax highlighting
- **Real-time Chat**: Streaming conversations with AI agents
- **Module Management**: Create, edit, and organize VBA modules
- **Context Awareness**: Integrates with active Excel workbook data

## Key Features

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

## Quick Start

See [SETUP.md](SETUP.md) for detailed setup instructions.

### Development
```bash
# Start backend
cd backend
python app.py

# Start frontend (new terminal)
cd frontend
npm start
```

### Production
```bash
# Build and deploy
scripts/deploy.bat
```

## API Documentation

See [API.md](API.md) for complete API documentation.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution guidelines.

## License

MIT License - see [LICENSE](../LICENSE) for details.