import { FolderIcon, ReturnIcon, CloseIcon } from './icons';

// Mock data for files/modules
const mockFiles = [
  { id: 1, name: 'Module 1', type: 'folder', tag: 'Personal' },
  { id: 2, name: 'Module 2', type: 'folder', tag: 'Personal' },
];

// Mock data for VBA shortcuts
const mockShortcuts = [
  { id: 1, name: 'Bottom bolded line', shortcut: 'Shift + w' },
  { id: 2, name: 'Apply Client Theme', shortcut: 'Shift + w' },
  { id: 3, name: 'Reload Pivot Tables', shortcut: 'Ctrl + Shift + w' },
  { id: 4, name: 'Bottom bolded line', shortcut: 'Ctrl + Alt + w' },
  { id: 5, name: 'Clean Data', shortcut: 'Shift + w' },
  { id: 6, name: 'Upload to Sharepoint', shortcut: 'Shift + w' },
  { id: 7, name: 'Auto-Fit & Zoom 100', shortcut: 'Ctrl + Shift + w' },
  { id: 8, name: 'Email as PDF', shortcut: 'Ctrl + Alt + w' },
];

const SearchMode = ({
  searchQuery,
  onSearchChange,
  onBuildModeClick,
  onFileClick,
  onShortcutClick,
  onClose,
}) => {
  // Filter files based on search
  const filteredFiles = mockFiles.filter((file) =>
    file.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  // Filter shortcuts based on search
  const filteredShortcuts = mockShortcuts.filter((shortcut) =>
    shortcut.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  return (
    <>
      {/* Header / Search Bar */}
      <header className="header">
        <div className="search-input-wrapper">
          <input
            type="text"
            className="search-input"
            placeholder="Search files and manage macros"
            value={searchQuery}
            onChange={(e) => onSearchChange(e.target.value)}
            autoFocus
          />
        </div>

        <div className="header-actions">
          <button className="build-mode-btn" onClick={onBuildModeClick}>
            AI Build Mode
            <span className="kbd">Tab</span>
          </button>
          <button className="close-btn" onClick={onClose}>
            <CloseIcon />
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="main-content">
        {/* All Files Section */}
        <div className="section-header">
          <span className="section-title">All Files</span>
          <span className="section-count">{filteredFiles.length} items</span>
        </div>

        <div className="file-list">
          {filteredFiles.map((file) => (
            <div
              key={file.id}
              className="file-item"
              onClick={() => onFileClick(file)}
            >
              <div className="file-icon">
                <FolderIcon size={24} />
              </div>
              <div className="file-info">
                <span className="file-name">
                  {file.name}
                  <span className="file-tag">{file.tag}</span>
                </span>
              </div>
              <span className="file-type">Macro Folder</span>
            </div>
          ))}
        </div>

        {/* VBA Shortcuts Section */}
        <div className="section-header">
          <span className="section-title">VBA Shortcuts</span>
          <span className="section-count">{filteredShortcuts.length} items</span>
        </div>

        <div className="shortcuts-grid">
          {filteredShortcuts.map((shortcut) => (
            <div
              key={shortcut.id}
              className="shortcut-item"
              onClick={() => onShortcutClick(shortcut)}
            >
              <span className="shortcut-icon">
                <ReturnIcon size={16} />
              </span>
              <span className="shortcut-name">{shortcut.name}</span>
              <span className="shortcut-keys">{shortcut.shortcut}</span>
            </div>
          ))}
        </div>
      </main>
    </>
  );
};

export default SearchMode;
