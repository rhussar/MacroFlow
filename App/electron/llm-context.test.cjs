const assert = require('node:assert/strict');
const test = require('node:test');

const { resolveWorkbookPromptContext, serializeWorkbookContext } = require('./llm-context');
const { resolveLlmPerformanceProfile } = require('./llm-performance');

function createMetadataResult() {
  return {
    success: true,
    workbookFound: true,
    llmContext: {
      structural: {
        workbookName: 'Book1.xlsm',
        worksheetNames: ['Sheet1', 'Sheet2', 'Sheet3'],
        activeSheet: 'Sheet1'
      },
      data: {
        usedRange: {
          address: 'A1:D12',
          rows: 12,
          columns: 4
        },
        headersByAddress: {
          A1: 'Customer',
          B1: 'Amount',
          C1: 'Date',
          D1: 'Status'
        },
        sampleRows: {
          rows: [
            ['Acme', 42, '2026-03-26', 'Open'],
            ['Globex', 64, '2026-03-27', 'Closed']
          ]
        },
        columns: [
          { column: 'A', header: 'Customer', typeSummary: 'text 2', examples: ['Acme', 'Globex'] },
          { column: 'B', header: 'Amount', typeSummary: 'number 2', examples: ['42', '64'] }
        ]
      },
      selection: {
        address: 'A1:D12',
        activeCell: {
          address: 'A1',
          value: 'Customer'
        },
        table: {
          name: 'CustomerTable',
          range: 'A1:D12'
        }
      }
    }
  };
}

test('resolveWorkbookPromptContext times out slow workbook metadata reads and falls back cleanly', async () => {
  const profile = resolveLlmPerformanceProfile({
    totalMemoryBytes: 8 * 1024 * 1024 * 1024,
    cpuCount: 4
  });

  const result = await resolveWorkbookPromptContext(
    {
      intent: 'create',
      workbookName: 'Book1.xlsm'
    },
    {
      performanceProfile: {
        ...profile,
        contextFetchTimeoutMs: 20
      },
      metadataByWorkbookImpl: () => new Promise(() => {})
    }
  );

  assert.equal(result.text, '');
  assert.equal(result.meta.timedOut, true);
  assert.equal(result.meta.skippedReason, 'timeout');
});

test('serializeWorkbookContext uses lighter context for low_resource profiles', () => {
  const lowResourceProfile = resolveLlmPerformanceProfile({
    totalMemoryBytes: 8 * 1024 * 1024 * 1024,
    cpuCount: 4
  });
  const standardProfile = resolveLlmPerformanceProfile({
    totalMemoryBytes: 32 * 1024 * 1024 * 1024,
    freeMemoryBytes: 20 * 1024 * 1024 * 1024,
    cpuCount: 12
  });

  const lowResourceText = serializeWorkbookContext(createMetadataResult(), {}, lowResourceProfile);
  const standardText = serializeWorkbookContext(createMetadataResult(), {}, standardProfile);

  assert.equal(lowResourceText.includes('Columns:'), false);
  assert.equal(standardText.includes('Columns:'), true);
  assert.ok(lowResourceText.length < standardText.length);
});
