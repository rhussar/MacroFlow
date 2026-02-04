import React, { useState } from 'react';
import {
  FolderIcon,
  FolderIconLarge,
  ReturnIcon,
  ArrowLeftIcon,
  CloseIcon,
  ListIcon,
  ChevronDownIcon,
} from './icons';
import CodePreview from './CodePreview';

// Mock data for all items (folders + macros)
const mockItems = [
  { id: 1, name: 'Module 1', type: 'folder', tag: 'Personal' },
  { id: 2, name: 'Module 2', type: 'folder', tag: 'Personal' },
  { id: 3, name: 'Bottom bolded line', type: 'macro', shortcut: 'Shift + w' },
  { id: 4, name: 'Reload Pivot Tables', type: 'macro', shortcut: 'Ctrl + Shift + w' },
  { id: 5, name: 'Clean Data', type: 'macro', shortcut: 'Shift + w' },
  { id: 6, name: 'Auto-Fit & Zoom 100', type: 'macro', shortcut: 'Ctrl + Shift + w' },
  { id: 7, name: 'Apply Client Theme', type: 'macro', shortcut: 'Shift + w' },
  { id: 8, name: 'Bottom bolded line', type: 'macro', shortcut: 'Ctrl + Alt + w' },
  { id: 9, name: 'Upload to Sharepoint', type: 'macro', shortcut: 'Shift + w' },
];

// Mock folder metadata
const mockFolderMetadata = {
  name: 'Module 1',
  type: 'VBA Module Folder',
  lastModified: 'Jan 25, 2026 (2h ago)',
  author: 'Ronan (Finance Ops)',
  contains: '5 functions, 3 macros',
  usage: '42 runs / week',
};

// Mock macro metadata
const mockMacroMetadata = {
  name: 'Reload Pivot Tables',
  type: 'VBA Macro',
  lastModified: 'Jan 20, 2026 (5d ago)',
  author: 'Personal',
  scope: 'Active Workbook',
  shortcut: 'Alt + Shift + C',
};

// Mock VBA code for macro preview
const mockMacroCode = `Sub ReloadPivotTables()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 1. Find Pivot Tables
    On Error Resume Next
    
    ws.Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    ' 2. Trim Whitespace
    For Each cell In ws.Range("B1:B150")
        cell.Value = Trim(cell.Value)
    Next cell
    
    ' 3. Fix Date Format
    ws.Columns("C:C").NumberFormat = "mm/dd/yyyy"
    
    MsgBox "Cleanup Complete!"
End Sub`;

const FileExplorer = ({ onBack, onClose }) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedItem, setSelectedItem] = useState(mockItems[3]); // Default to a macro

  // Filter items based on search
  const filteredItems = mockItems.filter((item) =>
    item.name.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const isFolder = selectedItem?.type === 'folder';
  const metadata = isFolder ? mockFolderMetadata : mockMacroMetadata;

  return (
    <>
      {/* Header */}
      <header className="header">
        <button className="header-back-btn" onClick={onBack}>
          <ArrowLeftIcon size={20} />
        </button>

        <div className="search-input-wrapper">
          <input
            type="text"
            className="search-input"
            placeholder="Search files..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
          />
        </div>

        <div className="header-actions">
          <button className="filter-dropdown">
            <ListIcon size={14} />
            All files
            <ChevronDownIcon size={14} />
          </button>
          <button className="close-btn" onClick={onClose}>
            <CloseIcon />
          </button>
        </div>
      </header>

      {/* Split View */}
      <div className="split-view">
        {/* Left Panel - File List */}
        <div className="split-left">
          <div className="section-header">
            <span className="section-title">All Files</span>
            <span className="section-count">{filteredItems.length} items</span>
          </div>

          <div className="file-list">
            {filteredItems.map((item) => (
              <div
                key={item.id}
                className={`file-item ${selectedItem?.id === item.id ? 'selected' : ''}`}
                onClick={() => setSelectedItem(item)}
              >
                <div className={`file-icon ${item.type === 'macro' ? 'macro' : ''}`}>
                  {item.type === 'folder' ? (
                    <FolderIcon size={20} />
                  ) : (
                    <ReturnIcon size={16} />
                  )}
                </div>
                <span className="file-name">{item.name}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Right Panel - Details */}
        <div className="split-right">
          {selectedItem && (
            <div className="details-panel">
              {isFolder ? (
                /* Folder View */
                <>
                  <div className="folder-icon-large">
                    <FolderIconLarge size={80} />
                  </div>

                  <div className="metadata-section">
                    <div className="metadata-title">Metadata</div>
                    <div className="metadata-row">
                      <span className="metadata-label">Name</span>
                      <span className="metadata-value">{metadata.name}</span>
                    </div>
                    <div className="metadata-row">
                      <span className="metadata-label">Type</span>
                      <span className="metadata-value">{metadata.type}</span>
                    </div>
                    <div className="metadata-row">
                      <span className="metadata-label">Last modified</span>
                      <span className="metadata-value">{metadata.lastModified}</span>
                    </div>
                    <div className="metadata-row">
                      <span className="metadata-label">Author</span>
                      <span className="metadata-value">{metadata.author}</span>
                    </div>
                    <div className="metadata-row">
                      <span className="metadata-label">Contains</span>
                      <span className="metadata-value">{metadata.contains}</span>
                    </div>
                    <div className="metadata-row">
                      <span className="metadata-label">Usage</span>
                      <span className="metadata-value">{metadata.usage}</span>
                    </div>
                  </div>
                </>
              ) : (
                /* Macro View */
                <>
                  <CodePreview
                    code={mockMacroCode}
                    title="Preview"
                    showHeader={true}
                    showEdit={true}
                    onEdit={() => console.log('Edit clicked')}
                  />

                  <div className="metadata-section">
                    <div className="metadata-title">Metadata</div>
                    <div className="metadata-row">
                      <span className="metadata-label">Name</span>
                      <span className="metadata-value">{metadata.name}</span>
                    </div>
                    <div className="metadata-row">
                      <span className="metadata-label">Type</span>
                      <span className="metadata-value">{metadata.type}</span>
                    </div>
                    <div className="metadata-row">
                      <span className="metadata-label">Last modified</span>
                      <span className="metadata-value">{metadata.lastModified}</span>
                    </div>
                    <div className="metadata-row">
                      <span className="metadata-label">Author</span>
                      <span className="metadata-value">{metadata.author}</span>
                    </div>
                    <div className="metadata-row">
                      <span className="metadata-label">Scope</span>
                      <span className="metadata-value">{metadata.scope}</span>
                    </div>
                    <div className="metadata-row">
                      <span className="metadata-label">Shortcut</span>
                      <span className="metadata-value">{metadata.shortcut}</span>
                    </div>
                  </div>
                </>
              )}
            </div>
          )}
        </div>
      </div>
    </>
  );
};

export default FileExplorer;