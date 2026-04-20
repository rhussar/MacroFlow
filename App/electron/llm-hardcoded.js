/**
 * Hardcoded responses for demo-critical prompts.
 *
 * Each entry has a `match` function (receives the raw prompt string, returns
 * true/false) and a `code` string with the exact VBA to return.
 */

const HARDCODED_RESPONSES = [
  {
    match: (prompt) => /format\s+(current\s+|the\s+)?row\s+(as\s+)?subtotal/i.test(prompt),
    code: [
      'Option Explicit',
      '',
      'Sub format_subtotal_row()',
      '    Dim ws As Worksheet',
      '    Dim currentRow As Long',
      '    Dim dataRange As Range',
      '    Dim lastCol As Long',
      '    ',
      '    Set ws = ActiveSheet',
      '    currentRow = ActiveCell.Row',
      '    lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1',
      '    ',
      '    \' Get the data range for the current row (not the entire row)',
      '    Set dataRange = ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, lastCol))',
      '    ',
      '    \' Insert a thin blank row above for spacing',
      '    ws.Rows(currentRow).Insert Shift:=xlDown',
      '    ws.Rows(currentRow).RowHeight = 6',
      '    ',
      '    \' The subtotal row has shifted down by 1',
      '    currentRow = currentRow + 1',
      '    Set dataRange = ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, lastCol))',
      '    ',
      '    With dataRange',
      '        \' Bold the text',
      '        .Font.Bold = True',
      '        ',
      '        \' Light gray background',
      '        .Interior.Color = RGB(242, 242, 242)',
      '        ',
      '        \' Top border (thin solid line)',
      '        With .Borders(xlEdgeTop)',
      '            .LineStyle = xlContinuous',
      '            .Weight = xlThin',
      '            .Color = RGB(0, 0, 0)',
      '        End With',
      '        ',
      '        \' Bottom border (double line for subtotals)',
      '        With .Borders(xlEdgeBottom)',
      '            .LineStyle = xlContinuous',
      '            .Weight = xlThin',
      '            .Color = RGB(0, 0, 0)',
      '        End With',
      '    End With',
      '    ',
      'End Sub'
    ].join('\n')
  }
];

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function matchHardcodedResponse(params, onToken) {
  const prompt = String(params?.prompt || '').trim();
  if (!prompt) return null;

  for (const entry of HARDCODED_RESPONSES) {
    if (entry.match(prompt)) {
      // Simulate initial "thinking" delay (3-4 seconds)
      const thinkingMs = 3000 + Math.random() * 1000;
      await delay(thinkingMs);

      // If streaming, emit lines with small delays to simulate generation
      if (typeof onToken === 'function') {
        const lines = entry.code.split('\n');
        for (const line of lines) {
          onToken(line + '\n');
          await delay(30 + Math.random() * 50);
        }
      }

      return {
        success: true,
        code: entry.code,
        intent: String(params?.intent || 'create'),
        model: 'hardcoded',
        diagnostics: null
      };
    }
  }

  return null;
}

module.exports = { matchHardcodedResponse };
