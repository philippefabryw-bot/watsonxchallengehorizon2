// ---------------------------------------------------------------------------
// watsonx.ai workback-plan proxy
// Runs server-side so the IBM Cloud API key never reaches the browser and
// the CORS problem disappears (frontend calls same-origin /api/* instead).
// ---------------------------------------------------------------------------
import express from 'express';

const app = express();
app.use(express.json({ limit: '1mb' }));

const PORT = process.env.PORT || 8080;
const API_KEY = process.env.WATSONX_API_KEY;
const DEPLOYMENT_URL =
  process.env.WATSONX_DEPLOYMENT_URL ||
  'https://au-syd.ml.cloud.ibm.com/ml/v1/deployments/wbplangen/text/generation?version=2021-05-01';
const STREAM_URL =
  process.env.WATSONX_STREAM_URL ||
  'https://au-syd.ml.cloud.ibm.com/ml/v1/deployments/wbplangen/text/generation_stream?version=2021-05-01';

if (!API_KEY) {
  console.error('FATAL: WATSONX_API_KEY environment variable is not set.');
  process.exit(1);
}

// ── IAM token cache ────────────────────────────────────────────────────────
// Tokens last ~1h; cache and refresh with a safety margin so we don't mint one
// on every request.
let cachedToken = null;
let cachedExpiry = 0;

async function getIamToken() {
  const now = Date.now();
  if (cachedToken && now < cachedExpiry - 60_000) {
    return cachedToken; // still valid (with 60s margin)
  }
  const resp = await fetch('https://iam.cloud.ibm.com/identity/token', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      Accept: 'application/json',
    },
    body: new URLSearchParams({
      grant_type: 'urn:ibm:params:oauth:grant-type:apikey',
      apikey: API_KEY,
    }),
  });
  if (!resp.ok) {
    const text = await resp.text().catch(() => resp.statusText);
    throw new Error(`IAM token request failed (${resp.status}): ${text}`);
  }
  const data = await resp.json();
  cachedToken = data.access_token;
  // expires_in is seconds from now
  cachedExpiry = now + (data.expires_in || 3600) * 1000;
  return cachedToken;
}

// ── Health check (used by OpenShift readiness/liveness probes) ──────────────
app.get('/api/health', (_req, res) => {
  res.json({ status: 'ok' });
});

// ── Non-streaming workback endpoint ────────────────────────────────────────
app.post('/api/workback', async (req, res) => {
  const usecase = (req.body?.usecase || '').toString().slice(0, 2000);
  if (!usecase) {
    return res.status(400).json({ error: 'Missing "usecase" in request body.' });
  }
  try {
    const token = await getIamToken();
    const wxResp = await fetch(DEPLOYMENT_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify({ parameters: { prompt_variables: { Usecase: usecase } } }),
    });
    const text = await wxResp.text();
    if (!wxResp.ok) {
      return res.status(wxResp.status).json({ error: `watsonx error ${wxResp.status}`, detail: text });
    }
    // Pass the watsonx JSON straight through
    res.type('application/json').send(text);
  } catch (err) {
    console.error('workback error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── Streaming workback endpoint (optional; token-by-token UX) ───────────────
// Proxies the watsonx SSE stream straight through to the browser.
app.post('/api/workback/stream', async (req, res) => {
  const usecase = (req.body?.usecase || '').toString().slice(0, 2000);
  if (!usecase) {
    return res.status(400).json({ error: 'Missing "usecase" in request body.' });
  }
  try {
    const token = await getIamToken();
    const wxResp = await fetch(STREAM_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Accept: 'text/event-stream',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify({ parameters: { prompt_variables: { Usecase: usecase } } }),
    });

    if (!wxResp.ok) {
      const text = await wxResp.text().catch(() => wxResp.statusText);
      return res.status(wxResp.status).json({ error: `watsonx error ${wxResp.status}`, detail: text });
    }

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');

    // Pipe the upstream SSE stream to the client
    const reader = wxResp.body.getReader();
    const decoder = new TextDecoder();
    // Abort upstream if the client disconnects
    req.on('close', () => reader.cancel().catch(() => {}));

    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      res.write(decoder.decode(value, { stream: true }));
    }
    res.end();
  } catch (err) {
    console.error('workback stream error:', err.message);
    if (!res.headersSent) res.status(500).json({ error: err.message });
    else res.end();
  }
});

app.listen(PORT, () => {
  console.log(`watsonx proxy listening on :${PORT}`);
});
