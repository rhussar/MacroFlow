const http = require('node:http');
const https = require('node:https');

function getTransport(url) {
  return String(url || '').trim().toLowerCase().startsWith('https:') ? https : http;
}

function requestText({ url, method = 'GET', headers = {}, body = '', timeoutMs = 30000 }) {
  return new Promise((resolve, reject) => {
    const transport = getTransport(url);
    const request = transport.request(url, { method, headers }, (response) => {
      let raw = '';
      response.setEncoding('utf8');
      response.on('data', (chunk) => {
        raw += chunk;
      });
      response.on('end', () => {
        resolve({
          statusCode: Number(response.statusCode) || 0,
          headers: response.headers || {},
          bodyText: raw
        });
      });
    });

    request.setTimeout(timeoutMs, () => {
      request.destroy(new Error('request timed out'));
    });

    request.on('error', (error) => {
      reject(error);
    });

    if (body) {
      request.write(body);
    }
    request.end();
  });
}

async function requestJson({
  url,
  method = 'GET',
  headers = {},
  bodyObject,
  timeoutMs = 30000
}) {
  const body = bodyObject === undefined ? '' : JSON.stringify(bodyObject);
  const response = await requestText({
    url,
    method,
    timeoutMs,
    headers: {
      Accept: 'application/json',
      ...(body
        ? {
            'Content-Type': 'application/json',
            'Content-Length': Buffer.byteLength(body)
          }
        : {}),
      ...headers
    },
    body
  });

  let parsedBody = null;
  try {
    parsedBody = response.bodyText ? JSON.parse(response.bodyText) : null;
  } catch {
    parsedBody = null;
  }

  return {
    ...response,
    parsedBody
  };
}

module.exports = {
  requestText,
  requestJson
};
