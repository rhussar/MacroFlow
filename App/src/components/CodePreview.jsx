import React, { useMemo, useRef, useEffect, useCallback } from 'react';

/**
 * Syntax highlight VBA code
 * Returns HTML with span elements for different code elements
 */
export const highlightVBA = (code) => {
  if (!code) return '';

  // Keywords (Sub, Function, Dim, Set, End, etc.)
  const keywords = /\b(Sub|End Sub|Function|End Function|Dim|Set|As|If|Then|Else|End If|For|To|Next|Each|In|Do|Loop|While|Wend|With|End With|Select|Case|End Select|On Error|Resume|GoTo|Exit|Private|Public|ByVal|ByRef|Optional|Const|Type|End Type|Enum|End Enum|Property|Get|Let|Nothing|New|Me|True|False|And|Or|Not|Mod|Is|Like)\b/g;

  // Data types
  const dataTypes = /\b(String|Integer|Long|Double|Single|Boolean|Variant|Object|Date|Currency|Byte|Worksheet|Workbook|Range|Collection|Dictionary)\b/g;

  // Built-in objects/functions
  const builtIns = /\b(MsgBox|InputBox|Debug|Print|ActiveSheet|ActiveWorkbook|ActiveCell|Application|ThisWorkbook|Cells|Columns|Rows|Range|Sheets|Worksheets|Workbooks)\b/g;

  // Process line by line to handle comments properly
  const lines = code.split('\n');
  const highlightedLines = lines.map((line) => {
    // Check if line has a comment
    const commentIndex = line.indexOf("'");

    if (commentIndex !== -1) {
      // Split into code and comment parts
      const codePart = line.substring(0, commentIndex);
      const commentPart = line.substring(commentIndex);

      // Highlight code part
      const highlightedCode = highlightCodePart(codePart, keywords, dataTypes, builtIns);

      // Wrap comment in span
      const highlightedComment = `<span class="code-comment">${escapeHtml(commentPart)}</span>`;

      return highlightedCode + highlightedComment;
    } else {
      return highlightCodePart(line, keywords, dataTypes, builtIns);
    }
  });

  return highlightedLines.join('\n');
};

const escapeHtml = (text) => {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
};

const highlightCodePart = (code, keywords, dataTypes, builtIns) => {
  let result = escapeHtml(code);

  // Highlight strings first (to avoid highlighting keywords inside strings)
  result = result.replace(/"([^"]*)"/g, '<span class="code-string">"$1"</span>');

  // Highlight numbers
  result = result.replace(/\b(\d+)\b/g, '<span class="code-number">$1</span>');

  // Highlight keywords
  result = result.replace(
    /\b(Sub|End Sub|Function|End Function|Dim|Set|As|If|Then|Else|End If|For|To|Next|Each|In|Do|Loop|While|Wend|With|End With|Select|Case|End Select|On Error|Resume|GoTo|Exit|Private|Public|ByVal|ByRef|Optional|Const|Type|End Type|Enum|End Enum|Property|Get|Let|Nothing|New|Me|True|False|And|Or|Not|Mod|Is|Like)\b/g,
    '<span class="code-keyword">$1</span>'
  );

  // Highlight data types
  result = result.replace(
    /\b(String|Integer|Long|Double|Single|Boolean|Variant|Object|Date|Currency|Byte|Worksheet|Workbook|Range|Collection|Dictionary)\b/g,
    '<span class="code-loop">$1</span>'
  );

  // Highlight built-ins
  result = result.replace(
    /\b(MsgBox|InputBox|Debug|Print|ActiveSheet|ActiveWorkbook|ActiveCell|Application|ThisWorkbook|Cells|Columns|Rows|Range|Sheets|Worksheets|Workbooks)\b/g,
    '<span class="code-function">$1</span>'
  );

  return result;
};

/**
 * CodePreview Component
 * Displays syntax-highlighted VBA code with optional actions
 */
const CodePreview = ({
  code,
  title = 'Preview',
  showHeader = true,
  editable = false,
  onChange,
  status = 'normal', // 'normal', 'success', 'error'
  errorLine = null,
  className = '',
}) => {
  const containerClass = `code-preview-container ${status} ${className}`.trim();
  const bodyRef = useRef(null);
  const isUserEditingRef = useRef(false);

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

  // Update innerHTML only when the code changes externally (not from user typing)
  useEffect(() => {
    if (!bodyRef.current) return;
    if (isUserEditingRef.current) {
      isUserEditingRef.current = false;
      return;
    }
    bodyRef.current.innerHTML = finalCode;
  }, [finalCode]);

  const handleInput = useCallback((e) => {
    if (!editable) return;
    isUserEditingRef.current = true;
    onChange?.(e.currentTarget.textContent);
  }, [editable, onChange]);

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
      />
    </div>
  );
};

export default CodePreview;