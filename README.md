# IBM Software Portfolio Advisor

An AI-powered account planning tool for IBM Technology Sellers. Upload a client's IBM software catalog (reference catalog XLSX), instantly get portfolio analysis, deployment gap insights, license modernization recommendations, and watsonx-generated workback plans — all in one single-container application deployable on OpenShift.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  Browser (React SPA)                                                         │
│                                                                              │
│  ┌──────────┐  ┌──────────────┐  ┌──────────────┐  ┌──────────────────────┐│
│  │ Overview │  │  Deployed /  │  │  Deployment  │  │  Licensing & Cost /  ││
│  │ Portfolio│  │  Available   │  │  Plan + LLM  │  │  Market Signals      ││
│  │ Snapshot │  │  Catalog     │  │  Workback    │  │  License Recs        ││
│  └──────────┘  └──────────────┘  └──────┬───────┘  └──────────────────────┘│
└─────────────────────────────────────────┼───────────────────────────────────┘
                                          │  POST /api/workback
                                          │  { usecase: "..." }
                            ┌─────────────▼───────────────────────────────┐
                            │  Express server  (server.js)  :8080          │
                            │                                               │
                            │  GET  /healthz        → liveness probe        │
                            │  GET  /api/health     → readiness probe       │
                            │  POST /api/workback   → watsonx proxy         │
                            │  GET  /*              → ./dist  (React SPA)  │
                            │                                               │
                            │  IAM token cache (server-side, ~1 h TTL)     │
                            └─────────────┬───────────────────────────────┘
                                          │  Bearer token (server-side only)
                                          │
                            ┌─────────────▼───────────────────────────────┐
                            │  IBM Cloud                                    │
                            │                                               │
                            │  IAM  https://iam.cloud.ibm.com              │
                            │    ↓  exchanges API key → access token        │
                            │  watsonx.ai AU-SYD                           │
                            │    /ml/v1/deployments/wbplangen/text/         │
                            │    generation  (LLM workback plan generator)  │
                            └─────────────────────────────────────────────┘
```

**Key design decisions:**

| Concern | Decision |
|---|---|
| API key security | Key lives in an OpenShift Secret → injected into the container as `WATSONX_API_KEY`. Never sent to the browser. |
| CORS | No CORS headers needed. The browser calls same-origin `/api/workback`; Express proxies to IBM Cloud server-side. |
| Deployment simplicity | One image, one process, no nginx, no sidecar. Express serves the Vite build from `./dist` **and** handles API routes. |
| Token management | IAM tokens are minted server-side, cached in memory for ~1 hour with a 60-second safety margin before expiry. |
| XLSX parsing | Runs entirely in the browser (no server upload). The `xlsx` library parses the file client-side; the data never leaves the device. |

---

## What it does

### Upload your client's catalog
Drag-and-drop or upload a **reference catalog XLSX** (the IBM Software Weighting/ACP pricing workbook). The app auto-detects the Weighting sheet, finds the real header row, and parses all line items. No data is sent to any server.

### Deployed vs. Available views
The catalog splits products into two views based on **Deployment Plan Quantity**:

- **Deployed tab** — products with `Deployment Plan Quantity > 0`. Shows deploy quantity, IBM cost per unit (*Drawdown Price – Discounted*), client price per unit (*Drawdown Price – Weighted*), extended IBM cost, and extended client value.
- **Available Catalog tab** — products with `Deployment Plan Quantity = 0`. Owned entitlements not yet deployed. Positioned as unrealised value opportunities for the sales conversation.

### Portfolio analytics
- **Overview** — snapshot stats: total items, deployed count, available count, total extended IBM cost, total extended client value, license model breakdown by category.
- **Deployment Plan** — use-case selector, effort/value matrix, AI workback plan generator.
- **Sales & Technical Plans** — structured business and technical plans for the selected use case.
- **Licensing & Cost** — full catalog pricing (IBM cost vs. client price), license modernization recommendations.
- **Market Signals** — AI/cloud market trends mapped to products in the client's catalog.
- **Client Presentations** — What / Why / How / Why Now narrative per product family.

### watsonx workback plan
Select a use case → click **Generate with watsonx** → the server calls the `wbplangen` LLM deployment and returns a multi-phase workback plan rendered as a table. The browser sends only `{ "usecase": "IBM Product — Use Case Title" }`. No token, no key, no CORS.

---

## Pricing field definitions

| Field in catalog | Column in XLSX | Meaning |
|---|---|---|
| **IBM cost** | `Drawdown Price – Discounted` | The discounted unit price IBM charges under the agreement |
| **Client price** | `Drawdown Price – Weighted` | The weighted unit price the client actually pays / will draw down |
| **Extended IBM cost** | `Extended Net Price` | IBM cost × deployment quantity |
| **Extended client value** | `Extended Weighted Price` | Client price × deployment quantity |
| **Deployed** | `Deployment Plan Quantity > 0` | Product is actively deployed |
| **Available** | `Deployment Plan Quantity = 0` | Entitlement owned, not yet deployed |

---

## Repository layout

```
├── src/
│   ├── App.jsx          # React SPA — all UI tabs, analysis engine, XLSX parser
│   ├── main.jsx         # React entry point
│   └── styles.css       # Dark-theme stylesheet
├── server.js            # Express: static serve + /api/workback proxy + health checks
├── package.json         # react + vite (build deps) + express (runtime dep)
├── vite.config.js       # Dev proxy: /api → http://localhost:8080
├── index.html           # Vite HTML entry
├── Dockerfile.openshift # Multi-stage: build React (UBI Node 20) → run Express
├── openshift/
│   ├── buildconfig.yaml    # S2I/Docker build from git repo
│   ├── imagestream.yaml    # Internal image stream
│   ├── deployment.yaml     # 2 replicas, WATSONX_API_KEY from secret
│   ├── service.yaml        # ClusterIP on :8080
│   ├── route.yaml          # Edge TLS route (HTTPS)
│   └── kustomization.yaml  # oc apply -k openshift/
└── reference catalog.xlsx  # Sample IBM software pricing workbook
```

---

## Local development

### Prerequisites
- Node.js ≥ 20
- An IBM Cloud API key with access to the `wbplangen` watsonx.ai deployment

### Option A — Hot-reload dev mode (recommended for UI work)

```bash
# Terminal 1: backend proxy
$env:WATSONX_API_KEY = "your-ibm-cloud-api-key"
node server.js
# Express listening on :8080

# Terminal 2: Vite dev server
npm run dev
# App at http://localhost:5173
# vite.config.js proxies /api/* → http://localhost:8080
```

### Option B — Full production stack locally

```bash
npm install
npm run build                              # produces ./dist
$env:WATSONX_API_KEY = "your-key"
node server.js                             # serves app + API on :8080
# open http://localhost:8080
```

---

## Deploy to OpenShift

### 1. Create the namespace (once)
```bash
oc new-project ibm-catalog-advisor
```

### 2. Create the watsonx API key secret (once)
```bash
oc create secret generic watsonx-api-key \
  --from-literal=apikey=YOUR_IBM_CLOUD_API_KEY \
  -n ibm-catalog-advisor
```

### 3. Apply all manifests
```bash
oc apply -k openshift/
```

### 4. Trigger a build
```bash
oc start-build ibm-catalog-advisor --follow -n ibm-catalog-advisor
```

### 5. Verify
```bash
oc get pods -n ibm-catalog-advisor
# NAME                                READY   STATUS    RESTARTS
# ibm-catalog-advisor-xxxx-yyyy       1/1     Running   0  (×2)

curl -k https://YOUR_ROUTE_HOST/healthz
# ok
```

The Deployment has an `image.openshift.io/triggers` annotation, so every successful build automatically triggers a rolling update — zero-downtime with `maxUnavailable: 0`.

---

## Environment variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `WATSONX_API_KEY` | Yes (prod) | — | IBM Cloud API key for watsonx.ai access |
| `WATSONX_DEPLOYMENT_URL` | No | `https://au-syd.ml.cloud.ibm.com/ml/v1/deployments/wbplangen/text/generation?version=2021-05-01` | Override if your deployment is in a different region or has a different ID |
| `PORT` | No | `8080` | HTTP port the server listens on |

---

## Tech stack

| Layer | Technology |
|---|---|
| Frontend | React 18, Vite 5 |
| XLSX parsing | SheetJS (`xlsx` 0.18) — client-side, no upload |
| Backend / API proxy | Node.js 20, Express 4 |
| AI | watsonx.ai (IBM Cloud AU-SYD), `wbplangen` prompt deployment |
| Container base | Red Hat UBI 9 Node.js 20 (`registry.access.redhat.com/ubi9/nodejs-20`) |
| Orchestration | Red Hat OpenShift (Kubernetes) |
| TLS | OpenShift edge termination — no certificate config in the app |
