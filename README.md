# IBM Software Portfolio Advisor

A single-container application: one Express server serves the built React app
**and** proxies watsonx.ai calls, so there's one image, one deployment, no CORS,
and the IBM Cloud API key never reaches the browser.

## Architecture

```
                 ┌────────────────────────────────────────┐
  browser ──────►│  Express (server.js) on :8080          │
                 │                                        │
                 │   /api/workback  → watsonx.ai (proxy)  │
                 │   /healthz       → probe               │
                 │   /* everything  → React app (./dist)  │
                 └────────────────────────────────────────┘
```

- `POST /api/workback` — mints an IAM token from `WATSONX_API_KEY` (server-side,
  cached ~1h) and calls the `wbplangen` deployment. The browser sends only
  `{ "usecase": "..." }` — no token, no key.
- Everything else is served from the Vite build output in `./dist`, with an
  SPA fallback to `index.html`.

## Local development

Two options:

**A. Full stack (matches production):**
```bash
npm install
npm run build                       # produces ./dist
WATSONX_API_KEY=your_key npm start  # serves app + api on :8080
# open http://localhost:8080
```

**B. Vite dev server with hot reload:**
```bash
npm install
WATSONX_API_KEY=your_key node server.js   # api on :8080 (terminal 1)
npm run dev                                # vite on :5173 (terminal 2)
# open http://localhost:5173  — vite.config.js proxies /api → :8080
```

## Deploy to OpenShift

### 1. Create the API-key secret (once)
```bash
oc create secret generic watsonx-api-key \
  --from-literal=apikey=YOUR_IBM_CLOUD_API_KEY \
  -n ibm-catalog-advisor
```

### 2. Apply the manifests
```bash
oc apply -k openshift/          # kustomize applies all resources
```

### 3. Build and roll out
```bash
oc start-build ibm-catalog-advisor --follow -n ibm-catalog-advisor
oc rollout status deployment/ibm-catalog-advisor -n ibm-catalog-advisor
```

The Deployment has an image trigger, so a new build auto-redeploys. Verify:
```bash
oc get pods -n ibm-catalog-advisor          # want 2/2 pods 1/1 READY
curl -k https://YOUR_ROUTE_HOST/healthz     # -> ok
```

## What lives where

```
├── server.js               # unified Express server (static + API)
├── package.json            # react + vite (build) + express (runtime)
├── Dockerfile.openshift     # multi-stage: build React, run Express
├── index.html / src/        # React app source (+ watsonx chat widget)
├── vite.config.js           # dev-server /api proxy
└── openshift/
    ├── buildconfig.yaml      # Docker build from Dockerfile.openshift
    ├── imagestream.yaml
    ├── deployment.yaml       # injects WATSONX_API_KEY from secret
    ├── service.yaml          # :8080
    ├── route.yaml            # edge TLS
    └── kustomization.yaml
```

## Key fixes applied

- **One app, one image.** Removed the separate `backend/` folder and the
  conflicting Dockerfiles; the frontend Dockerfile had been overwritten with
  backend content. Everything is now one Express process.
- **UBI npm EACCES fixed.** The Dockerfile copies only `package.json` (not the
  lockfile) and uses `--no-package-lock`, so npm never fails writing the
  lockfile under OpenShift's arbitrary UID.
- **No CORS, no browser token.** The API key stays server-side; the browser
  calls same-origin `/api/workback`.
- **Removed nginx.** Express serves the static build and the SPA fallback, so
  there's no nginx config to keep in sync.
