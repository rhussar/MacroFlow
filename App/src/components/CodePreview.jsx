import React, { useMemo, useRef, useEffect, useCallback } from 'react';

/**
 * Syntax highlight VBA code
 */
export const highlightVBA = (code) => {
  if (!code) return '';

  const keywords = /\b(Sub|End Sub|Function|End Function|Dim|Set|As|If|Then|Else|End If|For|To|Next|Each|In|Do|Loop|While|Wend|With|End With|Select|Case|End Select|On Error|Resume|GoTo|Exit|Private|Public|ByVal|ByRef|Optional|Const|Type|End Type|Enum|End Enum|Property|Get|Let|Nothing|New|Me|True|False|And|Or|Not|Mod|Is|Like)\b/g;
  const dataTypes = /\b(String|Integer|Long|Double|Single|Boolean|Variant|Object|Date|Currency|Byte|Worksheet|Workbook|Range|Collection|Dictionary)\b/g;
  const builtIns = /\b(MsgBox|InputBox|Debug|Print|ActiveSheet|ActiveWorkbook|ActiveCell|Application|ThisWorkbook|Cells|Columns|Rows|Range|Sheets|Worksheets|Workbooks)\b/g;

  const lines = code.split('\n');
  return lines.map((line) => {
    const commentIndex = line.indexOf("'");
    if (commentIndex !== -1) {
      return highlightCodePart(line.substring(0, commentIndex), keywords, dataTypes, builtIns) +
        `<span class="code-comment">${escapeHtml(line.substring(commentIndex))}</span>`;
    }
    return highlightCodePart(line, keywords, dataTypes, builtIns);
  }).join('\n');
};

const escapeHtml = (text) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const highlightCodePart = (code, keywords, dataTypes, builtIns) => {
  let result = escapeHtml(code);
  const strings = [];
  result = result.replace(/"[^"]*"/g, (match) => {
    const placeholder = `__STRING_${strings.length}__`;
    strings.push(`<span class="code-string">${match}</span>`);
    return placeholder;
  });
  result = result.replace(keywords, '<span class="code-keyword">$1</span>');
  result = result.replace(dataTypes, '<span class="code-type">$1</span>');
  result = result.replace(builtIns, '<span class="code-builtin">$1</span>');
  strings.forEach((str, i) => { result = result.replace(`__STRING_${i}__`, str); });
  return result;
};

const CodePreview = ({
  code,
  title = 'Preview',
  showHeader = true,
  editable = false,
  onChange,
  status = 'normal',
  errorLine = null,
  className = '',
}) => {
  const containerClass = `code-preview-container ${status} ${className}`.trim();
  const bodyRef = useRef(null);
  const hasFocusRef = useRef(false);

  const finalCode = useMemo(() => {
    const highlighted = highlightVBA(code);
    if (errorLine !== null && code) {
      const lines = highlighted.split('\n');
      if (lines[errorLine - 1]) {
        lines[errorLine - 1] = `<span class="code-line-error">${lines[errorLine - 1]}</span>`;
      }
      return lines.join('\n');
    }
    return highlighted;
  }, [code, errorLine]);

  // Apply highlighted HTML only when the editor does NOT have focus.
  // This covers: AI generation, initial load, navigating back, prop changes.
  // Never touch innerHTML while the user is in the editor — it destroys cursor.
  useEffect(() => {
    if (!bodyRef.current) return;
    if (hasFocusRef.current) return;
    bodyRef.current.innerHTML = finalCode;
  }, [finalCode]);

  const handleInput = useCallback((e) => {
    if (!editable) return;
    // Use innerText instead of textContent — it preserves line breaks from <div>/<br> elements
    onChange?.(e.currentTarget.innerText);
  }, [editable, onChange]);

  const handleKeyDown = useCallback((e) => {
    if (!editable) return;
    if (e.key === 'Tab') {
      e.preventDefault();
      document.execCommand('insertText', false, '    ');
    }
  }, [editable]);

  const handlePaste = useCallback((e) => {
    if (!editable) return;
    e.preventDefault();
    const pastedText = e.clipboardData.getData('text/plain');
    if (!pastedText) return;
    // Insert plain text at cursor — browser handles cursor position natively
    document.execCommand('insertText', false, pastedText);
  }, [editable]);

  const handleFocus = useCallback(() => {
    hasFocusRef.current = true;
  }, []);

  const handleBlur = useCallback(() => {
    hasFocusRef.current = false;
    // Re-apply syntax highlighting now that the user left the editor
    if (bodyRef.current) {
      bodyRef.current.innerHTML = finalCode;
    }
  }, [finalCode]);

  return (
    <div className={containerClass}>
      {showHeader && (
        <div className="code-preview-header">
          <span className="code-preview-title">{title}</span>
        </div>
      )}
      <div
        ref={bodyRef}
        className={`code-preview-body ${editable ? 'editable' : ''}`}
        contentEditable={editable}
        suppressContentEditableWarning={true}
        spellCheck={false}
        onInput={handleInput}
        onKeyDown={handleKeyDown}
        onPaste={handlePaste}
        onFocus={handleFocus}
        onBlur={handleBlur}
      />
    </div>
  );
};

export default CodePreview;
