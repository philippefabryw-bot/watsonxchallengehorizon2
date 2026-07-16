// ---------------------------------------------------------------------------
// IBM Software Portfolio Advisor — single-container server
//
// One Node/Express process does BOTH jobs:
//   1. Serves the built React app (static files from ./dist)
//   2. Proxies /api/workback to watsonx.ai, minting the IAM token server-side
//
// This means one image, one process, one deployment — no nginx, no CORS,
// and the IBM Cloud API key never reaches the browser.
// ---------------------------------------------------------------------------
import express from 'express';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const app = express();
app.use(express.json({ limit: '1mb' }));

const PORT = process.env.PORT || 8080;
const API_KEY = process.env.WATSONX_API_KEY;
const DEPLOYMENT_URL =
  process.env.WATSONX_DEPLOYMENT_URL ||
  'https://au-syd.ml.cloud.ibm.com/ml/v1/deployments/wbplangen/text/generation?version=2021-05-01';

// The static site lives next to this file, in ./dist (produced by `vite build`)
const STATIC_DIR = path.join(__dirname, 'dist');

// ---------------------------------------------------------------------------
// IAM token cache — tokens last ~1h; refresh with a safety margin.
// ---------------------------------------------------------------------------
let cachedToken = null;
let cachedExpiry = 0;

async function getIamToken() {
  const now = Date.now();
  if (cachedToken && now < cachedExpiry - 60_000) return cachedToken;

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
  cachedExpiry = now + (data.expires_in || 3600) * 1000;
  return cachedToken;
}

// ---------------------------------------------------------------------------
// Health check — used by OpenShift readiness/liveness probes
// ---------------------------------------------------------------------------
app.get('/healthz', (_req, res) => res.type('text/plain').send('ok\n'));
app.get('/api/health', (_req, res) => res.json({ status: 'ok' }));

// ---------------------------------------------------------------------------
// Workback generation — the one API endpoint the frontend calls
// ---------------------------------------------------------------------------
app.post('/api/workback', async (req, res) => {
  const usecase = (req.body?.usecase || '').toString().slice(0, 2000);
  if (!usecase) {
    return res.status(400).json({ error: 'Missing "usecase" in request body.' });
  }
  if (!API_KEY) {
    return res.status(500).json({ error: 'Server is missing WATSONX_API_KEY configuration.' });
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
    res.type('application/json').send(text);
  } catch (err) {
    console.error('workback error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ---------------------------------------------------------------------------
// Static React app + SPA fallback
// Serve built assets, and for any non-API route return index.html so client-
// side routing / deep links work.
// ---------------------------------------------------------------------------
app.use(express.static(STATIC_DIR));

app.get('*', (req, res, next) => {
  if (req.path.startsWith('/api/')) return next(); // don't swallow unknown API routes
  res.sendFile(path.join(STATIC_DIR, 'index.html'));
});

app.listen(PORT, () => {
  console.log(`IBM Portfolio Advisor listening on :${PORT}`);
  if (!API_KEY) {
    console.warn('WARNING: WATSONX_API_KEY is not set — /api/workback will return 500 until it is.');
  }
});
