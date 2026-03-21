const { MAX_WORKBOOK_CONTEXT_CHARS } = require('./llm-config');

const MAX_WORKSHEET_NAMES = 12;
const MAX_COLUMNS = 10;
const MAX_COLUMN_EXAMPLES = 3;
const MAX_SAMPLE_ROWS = 3;
const MAX_SAMPLE_ROW_COLUMNS = 6;
const MAX_SELECTION_VALUE_CHARS = 80;
const WORKBOOK_CONTEXT_BURST_CACHE_MS = 1000;

let cachedContextKey = '';
let cachedContextText = '';
let cachedContextAt = 0;
let cachedExcelBridge = null;

function toSafeString(value) {
  return String(value || '').trim();
}

function truncate(value, maxChars) {
  const normalized = String(value || '');
  if (normalized.length <= maxChars) {
    return normalized;
  }
  return normalized.slice(0, maxChars);
}

function getExcelBridge() {
  if (!cachedExcelBridge) {
    cachedExcelBridge = require('./excel-bridge');
  }
  return cachedExcelBridge;
}

function formatValue(value, maxChars = MAX_SELECTION_VALUE_CHARS) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  return truncate(String(value), maxChars);
}

function stringifyHeaders(headersByAddress) {
  if (!headersByAddress || typeof headersByAddress !== 'object') {
    return '';
  }

  return Object.entries(headersByAddress)
    .slice(0, MAX_COLUMNS)
    .map(([address, header]) => `${address}=${formatValue(header, 40) || '(blank)'}`)
    .join(', ');
}

function stringifyColumns(columns) {
  if (!Array.isArray(columns) || !columns.length) {
    return [];
  }

  return columns.slice(0, MAX_COLUMNS).map((column) => {
    const columnLabel = toSafeString(column?.column) || '?';
    const header = formatValue(column?.header, 50) || '(blank)';
    const typeSummary = formatValue(column?.typeSummary, 80) || 'unknown';
    const examples = Array.isArray(column?.examples)
      ? column.examples
          .filter((value) => value !== null && value !== undefined && value !== '')
          .slice(0, MAX_COLUMN_EXAMPLES)
          .map((value) => formatValue(value, 40))
          .filter(Boolean)
      : [];
    const parts = [`${columnLabel} "${header}"`, typeSummary];
    if (examples.length) {
      parts.push(`examples: ${examples.join(', ')}`);
    }
    return `- ${parts.join(' | ')}`;
  });
}

function stringifySampleRows(sampleRows, columns) {
  if (!Array.isArray(sampleRows) || !sampleRows.length) {
    return [];
  }

  const columnLabels = Array.isArray(columns)
    ? columns.slice(0, MAX_SAMPLE_ROW_COLUMNS).map((column, index) => {
        const header = formatValue(column?.header, 30);
        const fallback = toSafeString(column?.column) || `Col${index + 1}`;
        return header || fallback;
      })
    : [];

  return sampleRows.slice(0, MAX_SAMPLE_ROWS).map((row, rowIndex) => {
    const values = Array.isArray(row) ? row : [];
    const pairs = values.slice(0, MAX_SAMPLE_ROW_COLUMNS).map((value, valueIndex) => {
      const key = columnLabels[valueIndex] || `Col${valueIndex + 1}`;
      const normalizedValue = formatValue(value, 40) || '(blank)';
      return `${key}=${normalizedValue}`;
    });
    return `- Row ${rowIndex + 1}: ${pairs.join(', ')}`;
  });
}

function serializeWorkbookContext(metadataResult, options = {}) {
  const llmContext = metadataResult?.llmContext;
  if (!llmContext || typeof llmContext !== 'object') {
    return '';
  }

  const structural = llmContext.structural || {};
  const data = llmContext.data || {};
  const selection = llmContext.selection || {};
  const lines = [];

  const activeSheet = toSafeString(structural.activeSheet);
  if (activeSheet) {
    lines.push(`Active sheet: ${activeSheet}`);
  }

  const worksheetNames = Array.isArray(structural.worksheetNames)
    ? structural.worksheetNames.filter(Boolean).slice(0, MAX_WORKSHEET_NAMES)
    : [];
  if (worksheetNames.length) {
    lines.push(`Worksheets: ${worksheetNames.join(', ')}`);
  }

  const usedRange = data.usedRange || {};
  const usedRangeAddress = toSafeString(usedRange.address);
  if (usedRangeAddress) {
    const rowCount = Number(usedRange.rows) || 0;
    const columnCount = Number(usedRange.columns) || 0;
    lines.push(`Used range: ${usedRangeAddress} (${rowCount} rows x ${columnCount} columns)`);
  }

  const headersLine = stringifyHeaders(data.headersByAddress);
  if (headersLine) {
    lines.push(`Headers: ${headersLine}`);
  }

  const columnLines = stringifyColumns(data.columns);
  if (columnLines.length) {
    lines.push('Columns:');
    lines.push(...columnLines);
  }

  const sampleRowLines = stringifySampleRows(data.sampleRows?.rows, data.columns);
  if (sampleRowLines.length) {
    lines.push('Sample rows:');
    lines.push(...sampleRowLines);
  }

  const selectionAddress = toSafeString(selection.address);
  if (selectionAddress) {
    lines.push(`Selection: ${selectionAddress}`);
  }

  const activeCellAddress = toSafeString(selection?.activeCell?.address);
  if (activeCellAddress) {
    const activeCellValue = formatValue(selection?.activeCell?.value);
    lines.push(
      activeCellValue
        ? `Active cell: ${activeCellAddress} = ${activeCellValue}`
        : `Active cell: ${activeCellAddress}`
    );
  }

  const tableName = toSafeString(selection?.table?.name);
  const tableRange = toSafeString(selection?.table?.range);
  if (tableName) {
    lines.push(tableRange ? `Selected table: ${tableName} (${tableRange})` : `Selected table: ${tableName}`);
  }

  return truncate(lines.join('\n'), MAX_WORKBOOK_CONTEXT_CHARS);
}

function resolveWorkbookMetadata(args = {}, dependencies = {}) {
  const workbookName = toSafeString(args.workbookName);
  const workbookPath = toSafeString(args.workbookPath);
  const sheetName = toSafeString(args.sheetName);
  const metadataByWorkbookImpl =
    typeof dependencies.metadataByWorkbookImpl === 'function'
      ? dependencies.metadataByWorkbookImpl
      : (payload) => getExcelBridge().getWorksheetMetadataByWorkbookName(payload.workbookName, payload);
  const metadataImpl =
    typeof dependencies.metadataImpl === 'function'
      ? dependencies.metadataImpl
      : (payload) => getExcelBridge().getWorksheetMetadata(payload);

  if (workbookName || workbookPath) {
    return Promise.resolve(
      metadataByWorkbookImpl({
        workbookName,
        workbookPath,
        sheetName
      })
    );
  }

  return Promise.resolve(metadataImpl({ sheetName }));
}

async function resolveWorkbookPromptContext(args = {}, dependencies = {}) {
  const cacheKey = JSON.stringify({
    workbookName: toSafeString(args.workbookName),
    workbookPath: toSafeString(args.workbookPath),
    sheetName: toSafeString(args.sheetName),
    intent: toSafeString(args.intent)
  });
  const now = Date.now();
  if (
    cachedContextText &&
    cachedContextKey === cacheKey &&
    now - cachedContextAt < WORKBOOK_CONTEXT_BURST_CACHE_MS
  ) {
    return cachedContextText;
  }

  const metadataResult = await resolveWorkbookMetadata(args, dependencies);
  if (!metadataResult?.success) {
    return '';
  }
  if (metadataResult?.workbookFound === false) {
    return '';
  }
  const serialized = serializeWorkbookContext(metadataResult, args);
  cachedContextKey = cacheKey;
  cachedContextText = serialized;
  cachedContextAt = Date.now();
  return serialized;
}

module.exports = {
  serializeWorkbookContext,
  resolveWorkbookMetadata,
  resolveWorkbookPromptContext
};
