import React, { useEffect, useMemo, useState } from 'react';
import './App.css';

const DEFAULT_MACRO = `Sub HelloMacroFlow()
    MsgBox "Hello from MacroFlow!", vbInformation, "MacroFlow"
End Sub`;

const DEFAULT_MODULE = 'MacroFlowScratch';

function App() {
  const [connection, setConnection] = useState({
    available: false,
    workbooks: [],
    active: null,
    message: ''
  });
  const [modules, setModules] = useState([]);
  const [procedures, setProcedures] = useState([]);
  const [macroName, setMacroName] = useState('');
  const [moduleName, setModuleName] = useState(DEFAULT_MODULE);
  const [macroCode, setMacroCode] = useState(DEFAULT_MACRO);
  const [shortcutKey, setShortcutKey] = useState('');
  const [shortcutMessage, setShortcutMessage] = useState('');
  const [shortcutAudit, setShortcutAudit] = useState({ shortcuts: [], unmapped: [], note: '' });
  const [worksheets, setWorksheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [sheetMetadata, setSheetMetadata] = useState(null);
  const [logs, setLogs] = useState([]);
  const [pending, setPending] = useState({});

  const moduleStats = useMemo(() => {
    const stats = {};
    procedures.forEach((proc) => {
      if (!stats[proc.module]) {
        stats[proc.module] = { subs: 0, functions: 0, properties: 0 };
      }
      if (proc.kind.startsWith('Sub')) {
        stats[proc.module].subs += 1;
      } else if (proc.kind.startsWith('Function')) {
        stats[proc.module].functions += 1;
      } else {
        stats[proc.module].properties += 1;
      }
    });
    return stats;
  }, [procedures]);

  const setBusy = (key, value) => {
    setPending((prev) => ({ ...prev, [key]: value }));
  };

  const pushLog = (label, result) => {
    const entry = {
      time: new Date().toLocaleTimeString(),
      label,
      success: Boolean(result && result.success),
      message: result && result.message ? result.message : ''
    };
    setLogs((prev) => [entry, ...prev].slice(0, 8));
  };

  const withExcel = async (label, fn) => {
    if (typeof window === 'undefined' || !window.excel) {
      const result = { success: false, message: 'Excel bridge not available.' };
      pushLog(label, result);
      return result;
    }
    try {
      const result = await fn(window.excel);
      pushLog(label, result);
      return result;
    } catch (error) {
      const result = { success: false, message: error?.message || 'Unknown error.' };
      pushLog(label, result);
      return result;
    }
  };

  const refreshConnection = async () => {
    setBusy('refresh', true);
    const list = await withExcel('List workbooks', (excel) => excel.workbook.list());
    const info = await withExcel('Active workbook info', (excel) => excel.workbook.info());

    if (list.success) {
      setConnection({
        available: true,
        workbooks: list.workbooks || [],
        active: info.success ? info : null,
        message: ''
      });
      if (info.success && Array.isArray(info.sheets)) {
        const sheets = info.sheets.map((name, index) => ({ name, index: index + 1 }));
        setWorksheets((prev) => (prev.length ? prev : sheets));
        setSelectedSheet((prev) => {
          if (prev && info.sheets.includes(prev)) {
            return prev;
          }
          return info.sheets[0] || '';
        });
      }
    } else {
      setConnection({
        available: false,
        workbooks: [],
        active: null,
        message: list.message || 'Excel not found.'
      });
    }
    setBusy('refresh', false);
  };

  const loadModules = async () => {
    setBusy('modules', true);
    const result = await withExcel('List modules', (excel) => excel.vba.modules());
    if (result.success) {
      setModules(result.modules || []);
    } else {
      setModules([]);
    }
    setBusy('modules', false);
  };

  const loadProcedures = async () => {
    setBusy('procedures', true);
    const result = await withExcel('List procedures', (excel) => excel.vba.procedures());
    if (result.success) {
      setProcedures(result.procedures || []);
    } else {
      setProcedures([]);
    }
    setBusy('procedures', false);
  };

  const loadWorksheets = async () => {
    setBusy('sheets', true);
    const result = await withExcel('List worksheets', (excel) => excel.workbook.sheets());
    if (result.success) {
      setWorksheets(result.sheets || []);
      setSelectedSheet((prev) => {
        const names = (result.sheets || []).map((sheet) => sheet.name);
        if (prev && names.includes(prev)) {
          return prev;
        }
        return names[0] || '';
      });
    } else {
      setWorksheets([]);
    }
    setBusy('sheets', false);
  };

  const loadMetadata = async (sheetNameOverride) => {
    setBusy('metadata', true);
    const targetSheet = sheetNameOverride || selectedSheet;
    if (targetSheet) {
      setSelectedSheet(targetSheet);
    }
    const result = await withExcel('View worksheet metadata', (excel) =>
      excel.workbook.metadata({
        sheetName: targetSheet || undefined
      })
    );
    if (result.success) {
      setSheetMetadata(result);
    } else {
      setSheetMetadata(null);
    }
    setBusy('metadata', false);
  };

  const injectMacro = async () => {
    setBusy('inject', true);
    const result = await withExcel('Inject module', (excel) =>
      excel.vba.inject({ moduleName, code: macroCode })
    );
    if (result.success) {
      setMacroName('HelloMacroFlow');
    }
    setBusy('inject', false);
  };

  const runMacro = async () => {
    setBusy('run', true);
    await withExcel('Run macro', (excel) => excel.vba.run({ macroName }));
    setBusy('run', false);
  };

  const setShortcut = async () => {
    setBusy('shortcutSet', true);
    setShortcutMessage('');
    const result = await withExcel('Set shortcut', (excel) =>
      excel.vba.setShortcut({ macroName, shortcutKey })
    );
    setShortcutMessage(result.message || '');
    setBusy('shortcutSet', false);
  };

  const auditShortcuts = async () => {
    setBusy('shortcutAudit', true);
    const result = await withExcel('Audit shortcuts', (excel) => excel.vba.auditShortcuts());
    if (result.success) {
      setShortcutAudit({
        shortcuts: result.shortcuts || [],
        unmapped: result.unmapped || [],
        note: result.note || ''
      });
    } else {
      setShortcutAudit({ shortcuts: [], unmapped: [], note: result.message || '' });
    }
    setBusy('shortcutAudit', false);
  };

  useEffect(() => {
    refreshConnection();
  }, []);

  return (
    <div className="app-root">
      <header className="header">
        <div className="brand">
          <span className={`status-dot ${connection.available ? 'on' : 'off'}`} />
          <div>
            <div className="title">MacroFlow Bridge Lab</div>
            <div className="subtitle">COM bridge smoke tests</div>
          </div>
        </div>
        <div className="header-actions">
          <button className="btn" onClick={refreshConnection} disabled={pending.refresh}>
            {pending.refresh ? 'Refreshing...' : 'Refresh'}
          </button>
          <button className="btn ghost" onClick={() => window?.excel?.app?.close?.()}>
            Close
          </button>
        </div>
      </header>

      <main className="main-grid">
        <section className="panel">
          <div className="panel-header">
            <h2>Connection</h2>
            <span className={`chip ${connection.available ? 'ok' : 'warn'}`}>
              {connection.available ? 'Excel Detected' : 'Excel Missing'}
            </span>
          </div>
          <div className="panel-body">
            <div className="meta">
              <div className="meta-label">Active Workbook</div>
              <div className="meta-value">
                {connection.active ? connection.active.name : 'None'}
              </div>
              <div className="meta-sub">
                {connection.active ? connection.active.path : connection.message || 'Open Excel to connect.'}
              </div>
            </div>

            <div className="panel-section">
              <div className="panel-section-header">
                <span>Open Workbooks</span>
                <span className="count">{connection.workbooks.length}</span>
              </div>
              <div className="list">
                {connection.workbooks.length === 0 && (
                  <div className="empty">No open workbooks.</div>
                )}
                {connection.workbooks.map((wb) => (
                  <div className="list-item" key={wb.path}>
                    <div className="list-title">{wb.name}</div>
                    <div className="list-sub">{wb.path}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </section>

        <section className="panel">
          <div className="panel-header">
            <h2>Modules</h2>
            <button className="btn small" onClick={loadModules} disabled={pending.modules}>
              {pending.modules ? 'Scanning...' : 'Scan'}
            </button>
          </div>
          <div className="panel-body">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>Type</th>
                  <th>Lines</th>
                  <th>Subs</th>
                  <th>Funcs</th>
                </tr>
              </thead>
              <tbody>
                {modules.length === 0 && (
                  <tr>
                    <td colSpan="5" className="empty-cell">
                      No module data yet.
                    </td>
                  </tr>
                )}
                {modules.map((mod) => {
                  const stats = moduleStats[mod.name] || { subs: 0, functions: 0 };
                  return (
                    <tr key={mod.name}>
                      <td>{mod.name}</td>
                      <td>{mod.type}</td>
                      <td>{mod.lineCount}</td>
                      <td>{stats.subs}</td>
                      <td>{stats.functions}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </section>

        <section className="panel">
          <div className="panel-header">
            <h2>Macros</h2>
            <button className="btn small" onClick={loadProcedures} disabled={pending.procedures}>
              {pending.procedures ? 'Scanning...' : 'Scan'}
            </button>
          </div>
          <div className="panel-body">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>Kind</th>
                  <th>Scope</th>
                  <th>Module</th>
                </tr>
              </thead>
              <tbody>
                {procedures.length === 0 && (
                  <tr>
                    <td colSpan="4" className="empty-cell">
                      No macro data yet.
                    </td>
                  </tr>
                )}
                {procedures.map((proc, idx) => (
                  <tr
                    key={`${proc.module}-${proc.name}-${idx}`}
                    className="click-row"
                    onClick={() => setMacroName(`${proc.module}.${proc.name}`)}
                  >
                    <td>{proc.name}</td>
                    <td>{proc.kind}</td>
                    <td>{proc.scope}</td>
                    <td>{proc.module}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>

        <section className="panel span-2">
          <div className="panel-header">
            <h2>Worksheet Metadata</h2>
            <button className="btn small" onClick={loadWorksheets} disabled={pending.sheets}>
              {pending.sheets ? 'Loading...' : 'Load Worksheets'}
            </button>
          </div>
          <div className="panel-body">
            <div className="row">
              <div className="field">
                <label>How to Use</label>
                <div className="field-output">
                  Load worksheets, then click a row to view metadata and preview.
                </div>
              </div>
            </div>

            <div className="panel-section">
              <div className="panel-section-header">
                <span>Worksheets</span>
                <span className="count">{worksheets.length}</span>
              </div>
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>Rows</th>
                    <th>Cols</th>
                    <th>Used Range</th>
                  </tr>
                </thead>
                <tbody>
                  {worksheets.length === 0 && (
                    <tr>
                      <td colSpan="4" className="empty-cell">
                        No worksheet data loaded.
                      </td>
                    </tr>
                  )}
                  {worksheets.map((sheet) => (
                    <tr
                      key={sheet.name}
                      className={`click-row ${sheet.name === selectedSheet ? 'selected-row' : ''}`}
                      onClick={() => loadMetadata(sheet.name)}
                    >
                      <td>{sheet.name}</td>
                      <td>{sheet.usedRange ? sheet.usedRange.rows : '-'}</td>
                      <td>{sheet.usedRange ? sheet.usedRange.columns : '-'}</td>
                      <td>{sheet.usedRange ? sheet.usedRange.address : '-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="meta-grid">
              <div className="meta-card">
                <div className="meta-label">Workbook</div>
                <div className="meta-value">{sheetMetadata?.structuralContext?.workbook?.name || '--'}</div>
                <div className="meta-sub">
                  {sheetMetadata?.structuralContext?.workbook?.path || 'Open a workbook to view metadata.'}
                </div>
              </div>
              <div className="meta-card">
                <div className="meta-label">Active Sheet</div>
                <div className="meta-value">{sheetMetadata?.structuralContext?.activeSheet || '--'}</div>
                <div className="meta-sub">
                  {sheetMetadata?.dataContext?.usedRange
                    ? `${sheetMetadata.dataContext.usedRange.address} - ${sheetMetadata.dataContext.usedRange.rows} rows x ${sheetMetadata.dataContext.usedRange.columns} cols`
                    : 'Select a sheet and view metadata.'}
                </div>
              </div>
              <div className="meta-card">
                <div className="meta-label">Selection</div>
                <div className="meta-value">
                  {sheetMetadata?.selectionContext?.address || '--'}
                </div>
                <div className="meta-sub">
                  {sheetMetadata?.selectionContext?.inTable
                    ? `Table: ${sheetMetadata.selectionContext.table?.name}`
                    : 'Not in a table.'}
                </div>
              </div>
            </div>

            <div className="panel-section">
              <div className="panel-section-header">
                <span>Headers</span>
              </div>
              <div className="header-tags">
                {(sheetMetadata?.dataContext?.headers || []).length === 0 && (
                  <span className="empty">No headers captured.</span>
                )}
                {(sheetMetadata?.dataContext?.headers || []).map((header, idx) => (
                  <span className="tag" key={`${header}-${idx}`}>
                    {header || 'N/A'}
                  </span>
                ))}
              </div>
            </div>

            <div className="panel-section">
              <div className="panel-section-header">
                <span>Worksheet Tabs</span>
              </div>
              <div className="header-tags">
                {(sheetMetadata?.structuralContext?.sheets || []).length === 0 && (
                  <span className="empty">No sheet list.</span>
                )}
                {(sheetMetadata?.structuralContext?.sheets || []).map((sheet) => (
                  <span className="tag" key={sheet}>
                    {sheet}
                  </span>
                ))}
              </div>
            </div>

            <div className="panel-section">
              <div className="panel-section-header">
                <span>Column Summary (sampled)</span>
              </div>
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Col</th>
                    <th>Header</th>
                    <th>Type Mix</th>
                    <th>Non-Empty</th>
                    <th>Range</th>
                    <th>Examples</th>
                  </tr>
                </thead>
                <tbody>
                  {(sheetMetadata?.dataContext?.columns || []).length === 0 && (
                    <tr>
                      <td colSpan="6" className="empty-cell">
                        No column metadata yet.
                      </td>
                    </tr>
                  )}
                  {(sheetMetadata?.dataContext?.columns || []).map((col) => (
                    <tr key={`${col.column}-${col.index}`}>
                      <td>{col.column}</td>
                      <td>{col.header || 'N/A'}</td>
                      <td>{col.typeSummary}</td>
                      <td>{col.nonEmpty}</td>
                      <td>{col.numericMin !== null ? `${col.numericMin} -> ${col.numericMax}` : '-'}</td>
                      <td className="examples-cell">{col.examples.join(', ')}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="panel-section">
              <div className="panel-section-header">
                <span>Selection Context</span>
              </div>
              <div className="selection-grid">
                <div className="meta-card">
                  <div className="meta-label">Selection Address</div>
                  <div className="meta-value">{sheetMetadata?.selectionContext?.address || '--'}</div>
                </div>
                <div className="meta-card">
                  <div className="meta-label">Active Cell</div>
                  <div className="meta-value">{sheetMetadata?.selectionContext?.activeCell?.address || '--'}</div>
                  <div className="meta-sub">
                    {sheetMetadata?.selectionContext?.activeCell?.value ?? '--'}
                  </div>
                </div>
                <div className="meta-card">
                  <div className="meta-label">Table</div>
                  <div className="meta-value">
                    {sheetMetadata?.selectionContext?.inTable
                      ? sheetMetadata.selectionContext.table?.name
                      : '--'}
                  </div>
                  <div className="meta-sub">
                    {sheetMetadata?.selectionContext?.table?.range || 'Not inside a table.'}
                  </div>
                </div>
              </div>
            </div>

            <div className="panel-section">
              <div className="panel-section-header">
                <span>LLM Payload</span>
              </div>
              <pre className="code-block">
                {sheetMetadata?.llmContext
                  ? JSON.stringify(sheetMetadata.llmContext, null, 2)
                  : 'No payload yet.'}
              </pre>
            </div>

            <div className="panel-section">
              <div className="panel-section-header">
                <span>Preview</span>
              </div>
              {sheetMetadata?.preview?.rows && sheetMetadata.preview.rows.length > 0 ? (
                <table className="preview-table">
                  <tbody>
                    {sheetMetadata.preview.rows.map((row, rowIndex) => (
                      <tr key={`row-${rowIndex}`}>
                        {row.map((cell, cellIndex) => (
                          <td key={`cell-${rowIndex}-${cellIndex}`}>
                            {cell === null || cell === undefined ? '' : String(cell)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <div className="empty">No preview loaded.</div>
              )}
            </div>
          </div>
        </section>

        <section className="panel">
          <div className="panel-header">
            <h2>Macro Actions</h2>
          </div>
          <div className="panel-body">
            <div className="field">
              <label>Module Name</label>
              <input
                className="input"
                value={moduleName}
                onChange={(e) => setModuleName(e.target.value)}
              />
            </div>
            <div className="field">
              <label>Macro Code</label>
              <textarea
                className="textarea"
                value={macroCode}
                onChange={(e) => setMacroCode(e.target.value)}
              />
            </div>
            <div className="field-inline">
              <button className="btn" onClick={injectMacro} disabled={pending.inject}>
                {pending.inject ? 'Injecting...' : 'Inject Module'}
              </button>
              <input
                className="input"
                value={macroName}
                onChange={(e) => setMacroName(e.target.value)}
                placeholder="Macro name to run"
              />
              <button
                className="btn ghost"
                onClick={runMacro}
                disabled={pending.run || !macroName.trim()}
              >
                {pending.run ? 'Running...' : 'Run Macro'}
              </button>
            </div>
          </div>
        </section>

        <section className="panel">
          <div className="panel-header">
            <h2>Runtime + Shortcuts</h2>
            <button className="btn small" onClick={auditShortcuts} disabled={pending.shortcutAudit}>
              {pending.shortcutAudit ? 'Auditing...' : 'Audit Shortcuts'}
            </button>
          </div>
          <div className="panel-body">
            <div className="field">
              <label>Set Shortcut</label>
              <div className="field-inline">
                <input
                  className="input"
                  value={macroName}
                  onChange={(e) => setMacroName(e.target.value)}
                  placeholder="Macro name"
                />
                <input
                  className="input"
                  value={shortcutKey}
                  onChange={(e) => setShortcutKey(e.target.value)}
                  placeholder="Shortcut key (e.g., C)"
                />
                <button
                  className="btn small"
                  onClick={setShortcut}
                  disabled={pending.shortcutSet || !macroName.trim() || !shortcutKey.trim()}
                >
                  {pending.shortcutSet ? 'Saving...' : 'Set Shortcut'}
                </button>
              </div>
              {shortcutMessage && <div className="field-output">{shortcutMessage}</div>}
            </div>

            <div className="panel-section">
              <div className="panel-section-header">
                <span>Tracked Shortcuts</span>
              </div>
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Macro</th>
                    <th>Shortcut</th>
                  </tr>
                </thead>
                <tbody>
                  {shortcutAudit.shortcuts.length === 0 && (
                    <tr>
                      <td colSpan="2" className="empty-cell">
                        No tracked shortcuts yet.
                      </td>
                    </tr>
                  )}
                  {shortcutAudit.shortcuts.map((entry) => (
                    <tr key={`${entry.macro}-${entry.shortcut}`}>
                      <td>{entry.macro}</td>
                      <td>{entry.shortcut}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="panel-section">
              <div className="panel-section-header">
                <span>Unmapped Macros</span>
              </div>
              <div className="header-tags">
                {shortcutAudit.unmapped.length === 0 && (
                  <span className="empty">No unmapped macros.</span>
                )}
                {shortcutAudit.unmapped.map((macro) => (
                  <span className="tag" key={macro}>
                    {macro}
                  </span>
                ))}
              </div>
              {shortcutAudit.note && <div className="field-output">{shortcutAudit.note}</div>}
            </div>
          </div>
        </section>
      </main>

      <footer className="footer">
        <div className="footer-title">Recent Actions</div>
        <div className="log">
          {logs.length === 0 && <div className="empty">No actions yet.</div>}
          {logs.map((entry, idx) => (
            <div className={`log-row ${entry.success ? 'ok' : 'warn'}`} key={`${entry.time}-${idx}`}>
              <span className="log-time">{entry.time}</span>
              <span className="log-label">{entry.label}</span>
              <span className="log-message">{entry.message}</span>
            </div>
          ))}
        </div>
      </footer>
    </div>
  );
}

export default App;
