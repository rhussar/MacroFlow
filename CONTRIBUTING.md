# Contributing to MacroFlow

Thank you for your interest in contributing to MacroFlow! This document provides guidelines and information for contributors.

## 📋 Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [Project Structure](#project-structure)
- [Development Guidelines](#development-guidelines)
- [Submitting Changes](#submitting-changes)
- [Issue Guidelines](#issue-guidelines)
- [Testing](#testing)
- [Documentation](#documentation)

## Code of Conduct

This project adheres to a code of conduct that promotes a welcoming environment for all contributors. By participating, you are expected to uphold this code.

### Our Standards

- Be respectful and inclusive
- Focus on constructive feedback
- Help others learn and grow
- Maintain professional communication

## Getting Started

### Prerequisites

- Python 3.8+
- Node.js 16+
- Git
- Microsoft Excel (for testing)
- OpenAI API key

### Fork and Clone

1. Fork the repository on GitHub
2. Clone your fork locally:
   ```bash
   git clone https://github.com/YOUR_USERNAME/macroflow-ai.git
   cd macroflow-ai
   ```
3. Add upstream remote:
   ```bash
   git remote add upstream https://github.com/macroflow-ai/macroflow-ai.git
   ```

## Development Setup

1. **Environment Configuration:**
   ```bash
   cp .env.example .env
   # Edit .env with your OpenAI API key and configuration
   ```

2. **Backend Setup:**
   ```bash
   cd backend
   pip install -r requirements.txt
   ```

3. **Frontend Setup:**
   ```bash
   cd frontend
   npm install
   ```

4. **Start Development Servers:**
   ```bash
   # From project root
   scripts/start.bat    # Windows
   scripts/start.sh     # Linux/macOS
   ```

## Project Structure

```
MacroFlow/
├── backend/              # Python Flask backend
│   ├── agents/          # AI agent system
│   ├── app.py           # Main Flask application
│   └── config.py        # Configuration management
├── frontend/            # Office Add-in
│   ├── src/            # Source files
│   ├── assets/         # Static assets
│   └── manifest.xml    # Office Add-in manifest
├── docs/               # Documentation
├── scripts/            # Build and deployment scripts
├── config/             # Environment configurations
└── tests/              # Test files
```

## Development Guidelines

### Code Style

**Python (Backend):**
- Follow PEP 8 style guidelines
- Use descriptive variable and function names
- Add docstrings to all functions and classes
- Maximum line length: 100 characters
- Use type hints where appropriate

**JavaScript (Frontend):**
- Use ES6+ features
- Follow consistent naming conventions (camelCase)
- Add JSDoc comments for functions
- Use async/await for asynchronous operations
- Maximum line length: 100 characters

### Git Workflow

1. **Create a feature branch:**
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make your changes:**
   - Write clean, well-documented code
   - Follow existing code patterns
   - Add tests for new functionality

3. **Commit your changes:**
   ```bash
   git add .
   git commit -m "Add: Brief description of your changes"
   ```

   **Commit Message Format:**
   - `Add: New feature or functionality`
   - `Fix: Bug fixes`
   - `Update: Changes to existing functionality`
   - `Remove: Deletion of code or features`
   - `Docs: Documentation changes`
   - `Style: Code formatting changes`
   - `Test: Adding or updating tests`

4. **Push to your fork:**
   ```bash
   git push origin feature/your-feature-name
   ```

5. **Create a Pull Request**

### Pull Request Guidelines

1. **Before submitting:**
   - Ensure all tests pass
   - Update documentation if needed
   - Rebase your branch on the latest main branch
   - Ensure your code follows the style guidelines

2. **Pull Request Description:**
   - Clearly describe what your PR does
   - Reference any related issues
   - Include screenshots for UI changes
   - List any breaking changes

3. **PR Template:**
   ```markdown
   ## Description
   Brief description of changes

   ## Type of Change
   - [ ] Bug fix
   - [ ] New feature
   - [ ] Breaking change
   - [ ] Documentation update

   ## Testing
   - [ ] Tests pass locally
   - [ ] Manual testing completed
   - [ ] Added new tests for functionality

   ## Screenshots (if applicable)

   ## Related Issues
   Fixes #123
   ```

### Agent Development

When working on AI agents:

1. **Base Agent Pattern:**
   - Inherit from `BaseAgent` class
   - Implement required abstract methods
   - Add comprehensive docstrings

2. **Agent Testing:**
   - Test with various input scenarios
   - Verify error handling
   - Check response formatting

3. **Intent Classification:**
   - Update classifier when adding new intents
   - Test classification accuracy
   - Document intent patterns

### Frontend Development

**Office Add-in Guidelines:**
- Follow Microsoft's Office Add-in best practices
- Ensure HTTPS compatibility
- Test in multiple Excel versions if possible
- Validate manifest.xml changes

**UI/UX Considerations:**
- Maintain consistency with existing design
- Ensure accessibility compliance
- Test responsive behavior
- Follow Office Fluent UI patterns where applicable

## Testing

### Backend Testing

```bash
cd backend
python -m pytest tests/
```

**Test Guidelines:**
- Write unit tests for new functions
- Include integration tests for API endpoints
- Mock external dependencies (OpenAI API, Excel COM)
- Achieve minimum 80% code coverage

### Frontend Testing

```bash
cd frontend
npm test
npm run validate    # Validate Office manifest
```

**Test Guidelines:**
- Test Office.js integration
- Validate manifest structure
- Test error scenarios
- Check browser compatibility

### Manual Testing

1. **Complete User Flow:**
   - Start application
   - Open Excel with test workbook
   - Test VBA generation
   - Verify code injection
   - Test module management

2. **Error Scenarios:**
   - No Excel running
   - VBA access disabled
   - API key issues
   - Network connectivity problems

## Documentation

### Code Documentation

- Add docstrings to all Python functions
- Use JSDoc for JavaScript functions
- Comment complex logic
- Keep README files up to date

### API Documentation

When adding new endpoints:
1. Update `docs/API.md`
2. Include request/response examples
3. Document error scenarios
4. Add to Postman collection (if available)

### User Documentation

- Update setup instructions for new requirements
- Add troubleshooting sections for common issues
- Include screenshots for UI changes
- Keep changelog updated

## Issue Guidelines

### Reporting Bugs

**Use the bug report template:**
- Clear, descriptive title
- Steps to reproduce
- Expected vs actual behavior
- Environment details (OS, Excel version, etc.)
- Screenshots/logs if applicable

### Feature Requests

**Use the feature request template:**
- Clear description of the feature
- Use case and motivation
- Possible implementation approach
- Consider backward compatibility

### Security Issues

For security vulnerabilities:
- **DO NOT** create a public issue
- Email security@macroflow.ai
- Include detailed description and reproduction steps

## Release Process

### Version Numbers

We use Semantic Versioning (semver):
- `MAJOR.MINOR.PATCH`
- MAJOR: Breaking changes
- MINOR: New features, backward compatible
- PATCH: Bug fixes, backward compatible

### Release Checklist

1. Update version numbers
2. Update CHANGELOG.md
3. Run full test suite
4. Update documentation
5. Create release notes
6. Tag release in git
7. Deploy to staging
8. Deploy to production

## Community

### Getting Help

- **Documentation:** Check existing docs first
- **Discussions:** Use GitHub Discussions for questions
- **Discord:** Join our community Discord (link in README)
- **Email:** contact@macroflow.ai for general inquiries

### Recognition

Contributors are recognized in:
- CONTRIBUTORS.md file
- Release notes
- Annual contributor spotlight

## Development Tips

### Debugging

**Backend:**
```bash
FLASK_ENV=development python backend/app.py
```

**Frontend:**
- Use browser dev tools
- Check Office.js console
- Enable Office add-in debugging

### Performance

- Monitor API response times
- Optimize AI agent processing
- Cache frequently accessed data
- Profile memory usage

### Dependencies

- Keep dependencies up to date
- Check for security vulnerabilities
- Document any new dependencies
- Consider bundle size impact

## License

By contributing to MacroFlow, you agree that your contributions will be licensed under the MIT License.

---

**Thank you for contributing to MacroFlow! 🚀**

Your contributions help make VBA development more accessible and efficient for Excel users worldwide.