const assert = require('node:assert/strict');
const test = require('node:test');

const WindowFocusHelperClient = require('./window-focus-helper-client.js');

// Unit tests for resolve request lifecycle:
// - concurrent calls share one helper request
// - timeout path settles and clears in-flight state
function createClientWithFakeProcess() {
  const writes = [];
  const fakeProcess = {
    stdin: {
      destroyed: false,
      write: (chunk) => {
        writes.push(String(chunk));
      }
    },
    removeAllListeners: () => {},
    killed: false
  };

  const client = new WindowFocusHelperClient();
  client.process = fakeProcess;
  client.stopping = false;
  return { client, writes };
}

test('findExcelWithWorkbooks dedupes concurrent in-flight requests', async () => {
  const { client, writes } = createClientWithFakeProcess();

  const first = client.findExcelWithWorkbooks(250);
  const second = client.findExcelWithWorkbooks(250);

  assert.equal(first, second);
  assert.equal(writes.length, 1);
  assert.match(writes[0], /"type":"findExcelWithWorkbooks"/);

  client.handleMessage(
    JSON.stringify({
      type: 'excelResolved',
      found: true,
      activated: true,
      pid: 10840,
      workbookCount: 1,
      strategy: 'foreground'
    })
  );

  const [firstResult, secondResult] = await Promise.all([first, second]);
  assert.equal(firstResult.found, true);
  assert.equal(secondResult.found, true);
  assert.equal(firstResult.pid, 10840);
  assert.equal(client._pendingExcelResolve, null);
});

test('findExcelWithWorkbooks timeout resolves and clears pending promise state', async () => {
  const { client, writes } = createClientWithFakeProcess();

  const result = await client.findExcelWithWorkbooks(20);
  assert.equal(result.found, false);
  assert.equal(result.reason, 'timeout');
  assert.equal(writes.length, 1);
  assert.equal(client._pendingExcelResolve, null);
});
