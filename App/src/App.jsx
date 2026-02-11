import React, { useState, useEffect } from 'react';
import './App.css';

// Import components
import SearchMode from './components/SearchMode';
import BuildMode from './components/BuildMode';
import FileExplorer from './components/FileExplorer';
import ManualEditMode from './components/ManualEditMode';
import SettingsMenu from './components/SettingsMenu';
import { MacroFlowLogo } from './components/icons';

/**
 * Main App Component
 *
 * Modes:
 * - 'search': Main search window with files and VBA shortcuts
 * - 'build': AI Build mode for generating macros
 * - 'explorer': File explorer with details panel
 * - 'edit': Manual code editor mode
 */
function App() {
  // Current view mode
  const [mode, setMode] = useState('search');

  // Search query (shared between search and explorer)
  const [searchQuery, setSearchQuery] = useState('');

  // Settings menu open state
  const [settingsOpen, setSettingsOpen] = useState(false);

  // Handle keyboard shortcuts
  useEffect(() => {
    const handleKeyDown = (e) => {
      // Tab to toggle between Search and Build modes
      if (e.key === 'Tab' && !e.ctrlKey && !e.shiftKey && !e.altKey) {
        // Only toggle if not in an input field or if in our specific inputs
        const target = e.target;
        const isSearchInput = target.classList?.contains('search-input');

        if (isSearchInput || target.tagName !== 'INPUT') {
          e.preventDefault();
          setMode((prev) => {
            if (prev === 'search') return 'build';
            if (prev === 'build') return 'search';
            return prev;
          });
        }
      }

      // Alt+M for new macro (go to build mode)
      if (e.altKey && e.key === 'm') {
        e.preventDefault();
        setMode('build');
      }

      // Escape to close or go back
      if (e.key === 'Escape') {
        if (settingsOpen) {
          setSettingsOpen(false);
        } else if (mode === 'explorer' || mode === 'edit') {
          setMode('search');
        } else if (mode === 'build') {
          setMode('search');
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [mode, settingsOpen]);

  // Handle close app
  const handleClose = () => {
    // In Electron, this would close the window
    if (window.excel?.app?.close) {
      window.excel.app.close();
    } else {
      console.log('Close app');
    }
  };

  // Handle quit from settings
  const handleQuit = () => {
    handleClose();
  };

  // Handle file click (open explorer)
  const handleFileClick = (file) => {
    console.log('File clicked:', file);
    setMode('explorer');
  };

  // Handle shortcut click - runs the VBA macro
  const handleShortcutClick = async (shortcut) => {
    console.log('Shortcut clicked:', shortcut);

    // Try to run the macro via the Excel bridge
    if (window.excel?.vba?.run) {
      try {
        // Convert display name to macro name (e.g., "Auto-Fit & Zoom 100" -> "AutoFitAndZoom100")
        const macroName = shortcut.name
          .replace(/[^a-zA-Z0-9\s]/g, '') // Remove special chars
          .replace(/\s+/g, '_');           // Spaces to underscores

        console.log(`Running macro: ${macroName}`);
        const result = await window.excel.vba.run({ macroName });

        if (!result.success) {
          console.error('Macro execution failed:', result.message);
          // Could show a notification to the user here
        } else {
          console.log('Macro executed successfully:', result.message);
        }
      } catch (error) {
        console.error('Error running macro:', error);
      }
    } else {
      console.warn('Excel VBA API not available - running in browser mode?');
    }
  };

  // Render current mode content
  const renderContent = () => {
    switch (mode) {
      case 'search':
        return (
          <SearchMode
            searchQuery={searchQuery}
            onSearchChange={setSearchQuery}
            onBuildModeClick={() => setMode('build')}
            onFileClick={handleFileClick}
            onShortcutClick={handleShortcutClick}
            onClose={handleClose}
          />
        );

      case 'build':
        return (
          <BuildMode
            onBack={() => setMode('search')}
            onClose={handleClose}
            onEditMode={() => setMode('edit')}
          />
        );

      case 'explorer':
        return (
          <FileExplorer
            onBack={() => setMode('search')}
            onClose={handleClose}
          />
        );

      case 'edit':
        return (
          <ManualEditMode
            onBack={() => setMode('build')}
            onClose={handleClose}
          />
        );

      default:
        return null;
    }
  };

  // Check if we should show the default footer
  const showDefaultFooter = mode === 'search' || mode === 'explorer';

  return (
    <div className="app-container">
      {/* Main Content */}
      {renderContent()}

      {/* Default Footer (for search and explorer modes) */}
      {showDefaultFooter && (
        <footer className="footer">
          <div className="footer-left">
            <div
              className="logo"
              onClick={() => setSettingsOpen(!settingsOpen)}
            >
              <MacroFlowLogo size={20} />
            </div>
          </div>
          <div className="footer-right">
            <span className="footer-action">
              New macro
              <span className="kbd">Alt</span>
              <span className="kbd">M</span>
            </span>
          </div>
        </footer>
      )}

      {/* Settings Menu */}
      <SettingsMenu
        isOpen={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        onQuit={handleQuit}
      />
    </div>
  );
}

export default App;