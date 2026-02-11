import { getLlama, LlamaChatSession } from 'node-llama-cpp';

let session = null;
let model = null;
let context = null;

const systemPrompt = `You are MacroFlow AI, an expert Excel VBA developer.
- You write efficient, error-free VBA code.
- You ALWAYS add comments to explain complex logic.
- You never apologize, just write the code.
- If the user asks about non-Excel topics, politely refuse.
- Keep responses concise.`;

const handlers = {
  init: async ({ modelPath }) => {
    const llama = await getLlama();

    model = await llama.loadModel({
      modelPath,
      // gpuLayers: 0 // Uncomment to force CPU-only if needed
    });

    context = await model.createContext({
      contextSize: 4096,
    });

    session = new LlamaChatSession({
      contextSequence: context.getSequence(),
      systemPrompt,
    });

    return { success: true };
  },
  chat: async ({ prompt, context: excelContext = '' }) => {
    if (!session) {
      throw new Error('LLM not initialized. Call initLLM first.');
    }

    const fullPrompt = ` [CURRENT EXCEL CONTEXT]
${JSON.stringify(excelContext)}
[END CONTEXT]

USER REQUEST: ${prompt} `;

    const response = await session.prompt(fullPrompt);
    return { success: true, data: response };
  },
};

const send = (payload) => {
  process.stdout.write(`${JSON.stringify(payload)}\n`);
};

let buffer = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', (chunk) => {
  buffer += chunk;
  let index = buffer.indexOf('\n');
  while (index !== -1) {
    const line = buffer.slice(0, index).trim();
    buffer = buffer.slice(index + 1);
    if (line) {
      handleMessage(line);
    }
    index = buffer.indexOf('\n');
  }
});

async function handleMessage(line) {
  let message;
  try {
    message = JSON.parse(line);
  } catch (err) {
    send({ id: null, success: false, error: 'Invalid JSON message' });
    return;
  }

  const { id, type, ...payload } = message;
  const handler = handlers[type];
  if (!handler) {
    send({ id, success: false, error: `Unknown command: ${type}` });
    return;
  }

  try {
    const result = await handler(payload);
    send({ id, ...result });
  } catch (err) {
    send({ id, success: false, error: err?.message || String(err) });
  }
}
