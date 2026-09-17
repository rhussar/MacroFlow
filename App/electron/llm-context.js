const { MAX_WORKBOOK_CONTEXT_CHARS } = require('./llm-config');
const { resolveLlmPerformanceProfile } = require('./llm-performance');

const WORKBOOK_CONTEXT_BURST_CACHE_MS = 1000;
const MAX_SELECTION_VALUE_CHARS = 80;

const CONTEXT_MODE_LIMITS = {
  minimal: {
    maxWorksheetNames: 8,
    maxColumns: 0,
    maxColumnExamples: 0,
    maxSampleRows: 0,
    maxSampleRowColumns: 0,
    includeSelection: true
  },
  reduced: {
    maxWorksheetNames: 10,
    maxColumns: 6,
    maxColumnExamples: 1,
    maxSampleRows: 0,
    maxSampleRowColumns: 0,
    includeSelection: true
  },
  full: {
    maxWorksheetNames: 12,
    maxColumns: 10,
    maxColumnExamples: 3,
    maxSampleRows: 3,
    maxSampleRowColumns: 6,
    includeSelection: true
  }
};

let cachedExcelBridge = null;
let cachedContextEntry = null;

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

function withTimeout(promise, timeoutMs, errorMessage) {
  const normalizedTimeoutMs = Number(timeoutMs) || 0;
  if (normalizedTimeoutMs <= 0) {
    return Promise.resolve(promise);
  }

  let timeoutId = null;
  const timeoutPromise = new Promise((_, reject) => {
    timeoutId = setTimeout(() => {
      reject(new Error(errorMessage));
    }, normalizedTimeoutMs);
  });

  return Promise.race([Promise.resolve(promise), timeoutPromise]).finally(() => {
    if (timeoutId) {
      clearTimeout(timeoutId);
    }
  });
}

function stringifyHeaders(headersByAddress, maxColumns) {
  if (!headersByAddress || typeof headersByAddress !== 'object') {
    return '';
  }

  const normalizedMaxColumns = Number(maxColumns) || 0;
  return Object.entries(headersByAddress)
    .slice(0, normalizedMaxColumns > 0 ? normalizedMaxColumns : undefined)
    .map(([address, header]) => `${address}=${formatValue(header, 40) || '(blank)'}`)
    .join(', ');
}

function stringifyColumns(columns, modeLimits) {
  if (!Array.isArray(columns) || !columns.length || !modeLimits.maxColumns) {
    return [];
  }

  return columns.slice(0, modeLimits.maxColumns).map((column) => {
    const columnLabel = toSafeString(column?.column) || '?';
    const header = formatValue(column?.header, 50) || '(blank)';
    const typeSummary = formatValue(column?.typeSummary, 80) || 'unknown';
    const examples = Array.isArray(column?.examples)
      ? column.examples
          .filter((value) => value !== null && value !== undefined && value !== '')
          .slice(0, modeLimits.maxColumnExamples)
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

function stringifySampleRows(sampleRows, columns, modeLimits) {
  if (!Array.isArray(sampleRows) || !sampleRows.length || !modeLimits.maxSampleRows) {
    return [];
  }

  const columnLabels = Array.isArray(columns)
    ? columns.slice(0, modeLimits.maxSampleRowColumns).map((column, index) => {
        const header = formatValue(column?.header, 30);
        const fallback = toSafeString(column?.column) || `Col${index + 1}`;
        return header || fallback;
      })
    : [];

  return sampleRows.slice(0, modeLimits.maxSampleRows).map((row, rowIndex) => {
    const values = Array.isArray(row) ? row : [];
    const pairs = values.slice(0, modeLimits.maxSampleRowColumns).map((value, valueIndex) => {
      const key = columnLabels[valueIndex] || `Col${valueIndex + 1}`;
      const normalizedValue = formatValue(value, 40) || '(blank)';
      return `${key}=${normalizedValue}`;
    });
    return `- Row ${rowIndex + 1}: ${pairs.join(', ')}`;
  });
}

function resolveModeLimits(performanceProfile) {
  const mode = toSafeString(performanceProfile?.workbookContextMode).toLowerCase();
  return CONTEXT_MODE_LIMITS[mode] || CONTEXT_MODE_LIMITS.reduced;
}

function serializeWorkbookContext(
  metadataResult,
  options = {},
  performanceProfile = resolveLlmPerformanceProfile()
) {
  const llmContext = metadataResult?.llmContext;
  if (!llmContext || typeof llmContext !== 'object') {
    return '';
  }

  const modeLimits = resolveModeLimits(performanceProfile);
  const structural = llmContext.structural || {};
  const data = llmContext.data || {};
  const selection = llmContext.selection || {};
  const lines = [];

  const activeSheet = toSafeString(structural.activeSheet);
  if (activeSheet) {
    lines.push(`Active sheet: ${activeSheet}`);
  }

  const worksheetNames = Array.isArray(structural.worksheetNames)
    ? structural.worksheetNames.filter(Boolean).slice(0, modeLimits.maxWorksheetNames)
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

  const headersLine = stringifyHeaders(data.headersByAddress, Math.max(modeLimits.maxColumns, 8));
  if (headersLine) {
    lines.push(`Headers: ${headersLine}`);
  }

  const columnLines = stringifyColumns(data.columns, modeLimits);
  if (columnLines.length) {
    lines.push('Columns:');
    lines.push(...columnLines);
  }

  const sampleRowLines = stringifySampleRows(data.sampleRows?.rows, data.columns, modeLimits);
  if (sampleRowLines.length) {
    lines.push('Sample rows:');
    lines.push(...sampleRowLines);
  }

  if (modeLimits.includeSelection) {
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
  }

  const maxWorkbookContextChars =
    Number(performanceProfile?.maxWorkbookContextChars) || MAX_WORKBOOK_CONTEXT_CHARS;
  return truncate(lines.join('\n'), maxWorkbookContextChars);
}

function resolveWorkbookMetadata(args = {}, dependencies = {}) {
  const workbookName = toSafeString(args.workbookName);
  const workbookPath = toSafeString(args.workbookPath);
  const sheetName = toSafeString(args.sheetName);
  const contextMode = toSafeString(args.contextMode);
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
        sheetName,
        contextMode
      })
    );
  }

  return Promise.resolve(metadataImpl({ sheetName, contextMode }));
}

function buildContextMeta(overrides = {}) {
  return {
    included: false,
    chars: 0,
    cacheHit: false,
    timedOut: false,
    durationMs: 0,
    mode: 'reduced',
    skippedReason: '',
    ...overrides
  };
}

async function resolveWorkbookPromptContext(args = {}, dependencies = {}) {
  const performanceProfile = dependencies.performanceProfile || resolveLlmPerformanceProfile();
  const normalizedIntent = toSafeString(args.intent).toLowerCase();
  const contextMode = toSafeString(performanceProfile.workbookContextMode) || 'reduced';

  if (!performanceProfile.enableWorkbookContext) {
    return {
      text: '',
      meta: buildContextMeta({
        mode: contextMode,
        skippedReason: 'disabled'
      })
    };
  }

  if (normalizedIntent === 'ask' && !performanceProfile.includeWorkbookContextForAsk) {
    return {
      text: '',
      meta: buildContextMeta({
        mode: contextMode,
        skippedReason: 'ask_disabled'
      })
    };
  }

  const cacheKey = JSON.stringify({
    workbookName: toSafeString(args.workbookName),
    workbookPath: toSafeString(args.workbookPath),
    sheetName: toSafeString(args.sheetName),
    intent: normalizedIntent,
    mode: contextMode,
    maxWorkbookContextChars: Number(performanceProfile.maxWorkbookContextChars) || MAX_WORKBOOK_CONTEXT_CHARS
  });
  const now = Date.now();
  if (
    cachedContextEntry &&
    cachedContextEntry.key === cacheKey &&
    now - cachedContextEntry.at < WORKBOOK_CONTEXT_BURST_CACHE_MS
  ) {
    return {
      text: cachedContextEntry.text,
      meta: buildContextMeta({
        ...cachedContextEntry.meta,
        cacheHit: true,
        durationMs: 0
      })
    };
  }

  const startedAt = Date.now();
  try {
    const metadataResult = await withTimeout(
      resolveWorkbookMetadata({
        ...args,
        contextMode
      }, dependencies),
      performanceProfile.contextFetchTimeoutMs,
      'WORKBOOK_CONTEXT_TIMEOUT'
    );

    if (!metadataResult?.success) {
      return {
        text: '',
        meta: buildContextMeta({
          mode: contextMode,
          durationMs: Date.now() - startedAt,
          skippedReason: 'metadata_unavailable'
        })
      };
    }

    if (metadataResult?.workbookFound === false) {
      return {
        text: '',
        meta: buildContextMeta({
          mode: contextMode,
          durationMs: Date.now() - startedAt,
          skippedReason: 'workbook_not_found'
        })
      };
    }

    const text = serializeWorkbookContext(metadataResult, args, performanceProfile);
    const meta = buildContextMeta({
      included: Boolean(text),
      chars: text.length,
      mode: contextMode,
      durationMs: Date.now() - startedAt,
      skippedReason: text ? '' : 'empty'
    });

    cachedContextEntry = {
      key: cacheKey,
      text,
      meta,
      at: Date.now()
    };

    return {
      text,
      meta
    };
  } catch (error) {
    const timedOut = String(error?.message || '') === 'WORKBOOK_CONTEXT_TIMEOUT';
    return {
      text: '',
      meta: buildContextMeta({
        mode: contextMode,
        timedOut,
        durationMs: Date.now() - startedAt,
        skippedReason: timedOut ? 'timeout' : 'error'
      })
    };
  }
}

module.exports = {
  serializeWorkbookContext,
  resolveWorkbookMetadata,
  resolveWorkbookPromptContext
};
