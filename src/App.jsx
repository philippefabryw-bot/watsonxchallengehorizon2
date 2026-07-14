import { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

// ---------------------------------------------------------------------------
// Market signals – embedded knowledge base
// ---------------------------------------------------------------------------
const marketSignals = [
  {
    theme: 'AI governance and trusted AI',
    demand: 'High',
    summary:
      'Organizations are prioritizing governed GenAI rollouts, model risk controls, and explainability for regulated workloads.',
    drivers: ['EU AI Act readiness', 'Board-level risk oversight', 'Need for auditable copilots'],
    keywords: ['watsonx', 'openscale', 'spss', 'cognos', 'datastage'],
  },
  {
    theme: 'Application modernization and hybrid integration',
    demand: 'High',
    summary:
      'Enterprises continue moving legacy estates to container, API, and event-driven architectures while protecting core systems.',
    drivers: ['Cloud cost optimization', 'Faster release cycles', 'Mainframe and distributed coexistence'],
    keywords: ['cloud pak', 'urbancode', 'sterling', 'mq', 'api connect', 'app connect', 'integration'],
  },
  {
    theme: 'Data fabric and observability',
    demand: 'Medium',
    summary:
      'Data leaders are investing in governed data access, lineage, and monitoring to improve AI outcomes.',
    drivers: ['Need for trusted data products', 'Platform engineering adoption', 'Compliance reporting pressure'],
    keywords: ['datastage', 'information server', 'cognos', 'spss', 'network manager'],
  },
  {
    theme: 'Zero trust security modernization',
    demand: 'High',
    summary:
      'Security investments are shifting toward identity-aware controls, posture management, and automated response.',
    drivers: ['Ransomware pressure', 'Distributed workforce', 'Software supply chain exposure'],
    keywords: ['qradar', 'security verify', 'guardium', 'maas360', 'trusteer'],
  },
  {
    theme: 'DevSecOps and continuous delivery',
    demand: 'High',
    summary:
      'Development teams need integrated testing, deployment automation, and requirements traceability at enterprise scale.',
    drivers: ['Shift-left quality mandates', 'Audit trail requirements', 'Platform consolidation'],
    keywords: ['rational', 'devops', 'clearcase', 'clearquest', 'functional tester', 'test workbench', 'doors', 'engineering'],
  },
  {
    theme: 'Optimization and decision intelligence',
    demand: 'Medium',
    summary:
      'Firms are embedding mathematical optimization and simulation into operational processes to reduce cost and risk.',
    drivers: ['Supply chain resilience', 'Energy optimization targets', 'Regulatory capital efficiency'],
    keywords: ['cplex', 'ilog', 'optimization', 'decision'],
  },
];

// ---------------------------------------------------------------------------
// License modernization patterns derived from real license model vocabulary
// ---------------------------------------------------------------------------
const licenseModernizationRules = [
  {
    models: ['PVU / VPC'],
    title: 'PVU / VPC to subscription modernization',
    insight:
      'PVU and Virtual Processor Core licensing creates unpredictable cost as workloads scale. Migrate to subscription or consumption-aligned tiers so clients can right-size without renegotiating.',
  },
  {
    models: ['Authorized User', 'Floating User'],
    title: 'Seat-based to enterprise or consumption bundle',
    insight:
      'Authorized User and Floating User licenses become cost-inefficient as automation and API workloads grow. Consider enterprise agreements or API-call-based packages.',
  },
  {
    models: ['Session', 'Capacity'],
    title: 'Session / capacity pricing to cloud-native metering',
    insight:
      'Session and Establishment capacity metrics are hard to forecast in elastic environments. Align pricing to cloud-native unit economics (hours, API calls, or resource units).',
  },
  {
    models: ['Per Server'],
    title: 'Per-server license to virtual-fleet or container-aware model',
    insight:
      'Per-virtual-server licenses do not scale economically with container orchestration. Transition to node-pool or managed-server-count pricing.',
  },
];

// ---------------------------------------------------------------------------
// Helpers: infer product family and license model from raw Description string
// ---------------------------------------------------------------------------
const deriveProductFamily = (desc) => {
  if (!desc) return 'Unknown';
  const m = desc.match(
    /^IBM\s+([\w\s\-+]+?)\s+(?:Cartridge|Authorized|Floating|Virtual|Non-Production|Processor|Resource|Managed|Software|Developer|Simultaneous|Per\s|per\s|Suite|High|Data\s|Build\s|Deploy\s|Velocity|Committer)/i,
  );
  if (m) return `IBM ${m[1].trim()}`;
  // Fallback: first 5 words
  return desc.split(/\s+/).slice(0, 5).join(' ');
};

const deriveLicenseModel = (desc) => {
  const d = (desc || '').toLowerCase();
  if (d.includes('authorized user')) return 'Authorized User';
  if (d.includes('floating user')) return 'Floating User';
  if (d.includes('virtual processor core')) return 'PVU / VPC';
  if (d.includes('processor value unit') || d.includes('(pvu)')) return 'PVU / VPC';
  if (d.includes('resource unit')) return 'Resource Unit';
  if (d.includes('per virtual server') || d.includes('managed') && d.includes('server')) return 'Per Server';
  if (d.includes('simultaneous session')) return 'Session';
  if (d.includes('per process') || d.includes('per establishment')) return 'Capacity';
  if (d.includes('managed device')) return 'Per Device';
  return 'Other';
};

const deriveCategory = (desc) => {
  const d = (desc || '').toLowerCase();
  if (d.includes('qradar') || d.includes('security') || d.includes('trusteer') || d.includes('guardium') || d.includes('maas360')) return 'Security';
  if (d.includes('cognos') || d.includes('spss') || d.includes('datastage') || d.includes('information server')) return 'Data & Analytics';
  if (d.includes('devops') || d.includes('rational') || d.includes('clearcase') || d.includes('clearquest') || d.includes('functional tester') || d.includes('test workbench') || d.includes('doors') || d.includes('engineering')) return 'DevOps & QA';
  if (d.includes('cplex') || d.includes('ilog') || d.includes('optimization')) return 'Optimization';
  if (d.includes('cloud pak') || d.includes('urbancode') || d.includes('sterling') || d.includes('mq') || d.includes('integration') || d.includes('api')) return 'Integration';
  if (d.includes('application runtime') || d.includes('websphere')) return 'Application Runtime';
  return 'Other';
};

// ---------------------------------------------------------------------------
// XLSX parsing: reads all rows from the first sheet that has the ACP columns
// ---------------------------------------------------------------------------
const ACP_HEADERS = [
  'S&S part number',
  'License part number',
  'Description',
  'Royalty',
  'Consumption rate (CAD)',
  'ACP S&S Unit Reference Price for NTE for 2029 (CAD)',
  'ACP S&S Unit Reference Price for NTE for 2030 (CAD)',
];

const parseXlsxWorkbook = (workbook) => {
  // Try to find the ACP sheet; fall back to first sheet
  const sheetName =
    workbook.SheetNames.find((n) => /ACP/i.test(n)) ?? workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Convert to array of arrays to find the real header row
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

  // Find the row index where ACP_HEADERS[0] appears (skip title rows)
  const headerRowIndex = raw.findIndex((row) =>
    row.some((cell) => typeof cell === 'string' && cell.trim() === ACP_HEADERS[2]),
  );
  if (headerRowIndex < 0) return [];

  const headers = raw[headerRowIndex].map((h) => (h ? String(h).trim() : ''));
  const descIdx = headers.findIndex((h) => h === 'Description');
  const royaltyIdx = headers.findIndex((h) => h === 'Royalty');
  const consumptionIdx = headers.findIndex((h) => /Consumption rate/i.test(h));
  const price2029Idx = headers.findIndex((h) => /2029/i.test(h));
  const price2030Idx = headers.findIndex((h) => /2030/i.test(h));
  const ssPartIdx = headers.findIndex((h) => /S&S part/i.test(h));
  const licPartIdx = headers.findIndex((h) => /License part/i.test(h));

  return raw
    .slice(headerRowIndex + 1)
    .filter((row) => row[descIdx] && typeof row[descIdx] === 'string' && row[descIdx].trim().startsWith('IBM'))
    .map((row, index) => {
      const desc = String(row[descIdx] ?? '').trim();
      return {
        id: `row-${index}`,
        ssPart: String(row[ssPartIdx] ?? '').trim(),
        licensePart: String(row[licPartIdx] ?? '').trim(),
        description: desc,
        royalty: String(row[royaltyIdx] ?? '').trim(),
        consumptionRate: Number(row[consumptionIdx]) || null,
        price2029: Number(row[price2029Idx]) || null,
        price2030: Number(row[price2030Idx]) || null,
        productFamily: deriveProductFamily(desc),
        licenseModel: deriveLicenseModel(desc),
        category: deriveCategory(desc),
      };
    });
};

// ---------------------------------------------------------------------------
// Built-in sample catalog seeded from the real XLSX data
// ---------------------------------------------------------------------------
const SAMPLE_ENTRIES = [
  { desc: 'IBM Cognos Analytics Administrator Cartridge for Authorized User License + SW Subscription and Support 12 Months', royalty: 'Yes', consumption: 19498.48, p2029: 3417.94, p2030: 3657.2, ssPart: 'E03GSZX', licPart: 'D03GTZX' },
  { desc: 'IBM Cognos Analytics Explorer Cartridge for Authorized User License + SW Subscription and Support 12 Months', royalty: 'Yes', consumption: 3470.39, p2029: 607.42, p2030: 649.94, ssPart: 'E03H8ZX', licPart: 'D03H9ZX' },
  { desc: 'IBM DataStage Enterprise Plus Cartridge for Virtual Processor Core License + SW Subscription and Support 12 Months', royalty: 'No', consumption: 73656.92, p2029: 12890.49, p2030: 13792.82, ssPart: 'E03JRZX', licPart: 'D03JSZX' },
  { desc: 'IBM DataStage Enterprise Plus Cartridge for Non-Production Virtual Processor Core License + SW Subscription and Support 12 Months', royalty: 'No', consumption: 36776.79, p2029: 6444.7, p2030: 6895.83, ssPart: 'E03JZZX', licPart: 'D03K0ZX' },
  { desc: 'IBM Security QRadar Suite Software 100 Resource Unit License + Subscription & Support 12 Months', royalty: 'Yes', consumption: 3869.54, p2029: 675.33, p2030: 722.6, ssPart: 'E0AE3ZX', licPart: 'D0AE4ZX' },
  { desc: 'IBM Security QRadar Software Install License + SW Subscription & Support 12 Months', royalty: 'No', consumption: 10299.84, p2029: 1815.73, p2030: 1942.83, ssPart: 'E0NBALL', licPart: 'D1RNCLL' },
  { desc: 'IBM Security QRadar Event Capacity 2.5K Events Per Second License + SW Subscription & Support 12 Months', royalty: 'No', consumption: 115140.14, p2029: 20017.59, p2030: 21418.82, ssPart: 'E0NBILL', licPart: 'D1RP3LL' },
  { desc: 'IBM SPSS Statistics Standard Authorized User License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 7963.04, p2029: 1393.58, p2030: 1491.13, ssPart: 'E09PILL', licPart: 'D0EKZLL' },
  { desc: 'IBM Network Manager Managed Device License + SW Subscription and Support 12 Months', royalty: 'No', consumption: 141.49, p2029: 24.73, p2030: 26.46, ssPart: 'E0R56LL', licPart: 'D28V7LL' },
  { desc: 'IBM ILOG CPLEX Optimizer Cluster Edition per Process License + SW Subscription & Support 12 Months', royalty: 'No', consumption: 8865.57, p2029: 1557.12, p2030: 1666.12, ssPart: 'E0N6PLL', licPart: 'D1QZVLL' },
  { desc: 'IBM ILOG CPLEX Optimization Studio Developer Edition Floating User Single Session License + SW Subscription & Support 12 Months', royalty: 'No', consumption: 19053.65, p2029: 3340.25, p2030: 3574.07, ssPart: 'E0JZWLL', licPart: 'D16I8LL' },
  { desc: 'IBM Enterprise Application Runtimes Virtual Processor Core License + Software Subscription & Support 12 Months', royalty: 'Yes', consumption: 21247.2, p2029: 3718.39, p2030: 3978.68, ssPart: 'E0ZNZZX', licPart: 'D0ZP1ZX' },
  { desc: 'IBM DevOps Build Committer Per Floating User Single Install License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 837.33, p2029: 146.15, p2030: 156.38, ssPart: 'E0KDQLL', licPart: 'D1ABHLL' },
  { desc: 'IBM DevOps Velocity User per Floating User License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 21341.22, p2029: 3729.26, p2030: 3990.31, ssPart: 'E0PDALL', licPart: 'D1YIILL' },
  { desc: 'IBM Rational Application Developer for WebSphere Software Floating User Single Install License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 12659.71, p2029: 2226.47, p2030: 2382.32, ssPart: 'E01MKLL', licPart: 'D54NBLL' },
  { desc: 'IBM DevOps Code ClearCase Floating User License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 7556.35, p2029: 1331.1, p2030: 1424.28, ssPart: 'E012TLL', licPart: 'D5315LL' },
  { desc: 'IBM Rational ClearQuest Authorized User License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 2960.81, p2029: 517.23, p2030: 553.44, ssPart: 'E019BLL', licPart: 'D53NSLL' },
  { desc: 'IBM Rational Functional Tester Floating User License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 17908.97, p2029: 3145.2, p2030: 3365.36, ssPart: 'E012DLL', licPart: 'D530BLL' },
  { desc: 'IBM DevOps Test Performance Virtual Users 1000 Floating User License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 109439.39, p2029: 19152.64, p2030: 20493.32, ssPart: 'E012LLL', licPart: 'D530RLL' },
  { desc: 'IBM Rational Test Virtualization Server Processor Value Unit (PVU) License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 3194.09, p2029: 559.06, p2030: 598.19, ssPart: 'E0DU6LL', licPart: 'D0PTZLL' },
  { desc: 'IBM Rational Test Workbench Authorized User Single Install License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 17617.15, p2029: 3083.26, p2030: 3299.09, ssPart: 'E0DTRLL', licPart: 'D0PS7LL' },
  { desc: 'IBM UrbanCode Deploy Server Agent per Virtual Server License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 144774.81, p2029: 25225.18, p2030: 26990.94, ssPart: 'E0J1LLL', licPart: 'D12U7LL' },
  { desc: 'IBM UrbanCode Velocity for Authorized User Single Install License + SW Subscription and Support 12 Months', royalty: 'Yes', consumption: 5141.05, p2029: 902.98, p2030: 966.19, ssPart: 'E0R1MLL', licPart: 'D28BMLL' },
  { desc: 'IBM Engineering Requirements Management DOORS family Authorized User License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 6616.13, p2029: 1152.35, p2030: 1232.01, ssPart: 'E074XLL', licPart: 'D09KNLL' },
  { desc: 'IBM Cloud Pak for Integration Virtual Processor Core License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 24788.55, p2029: 4338.3, p2030: 4641.98, ssPart: 'E0PPBLL', licPart: 'D20ZBLL' },
  { desc: 'IBM Sterling B2B Integrator Basic Edition for Non-Production Environment Processor Value Unit (PVU) License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 140.5, p2029: 24.67, p2030: 26.39, ssPart: 'E0J0ELL', licPart: 'D12RMLL' },
  { desc: 'IBM Information Server Cartridge for Virtual Processor Core License + SW Subscription & Support 12 Months', royalty: 'No', consumption: 117669.86, p2029: 20637.5, p2030: 22082.13, ssPart: 'E034QZX', licPart: 'D034RZX' },
  { desc: 'IBM QRadar High Availability Software Install License + SW Subscription & Support 12 Months', royalty: 'No', consumption: 8266.84, p2029: 1438.13, p2030: 1538.8, ssPart: 'E0NCHLL', licPart: 'D1RS0LL' },
  { desc: 'IBM ILOG CPLEX Optimizer Cluster Edition per Establishment License + SW Subscription & Support 12 Months', royalty: 'No', consumption: 118082.32, p2029: 20553.83, p2030: 21992.6, ssPart: 'E0N6ZLL', licPart: 'D1R0PLL' },
  { desc: 'IBM Rational ClearQuest Floating User License + SW Subscription & Support 12 Months', royalty: 'Yes', consumption: 8243.33, p2029: 1432.7, p2030: 1532.99, ssPart: 'E0130LL', licPart: 'D531LLL' },
].map((e, i) => ({
  id: `sample-${i}`,
  ssPart: e.ssPart,
  licensePart: e.licPart,
  description: e.desc,
  royalty: e.royalty,
  consumptionRate: e.consumption,
  price2029: e.p2029,
  price2030: e.p2030,
  productFamily: deriveProductFamily(e.desc),
  licenseModel: deriveLicenseModel(e.desc),
  category: deriveCategory(e.desc),
}));

// ---------------------------------------------------------------------------
// Analysis engine
// ---------------------------------------------------------------------------
const fmt = (n) =>
  n == null ? '—' : `CA$${Number(n).toLocaleString('en-CA', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

const analyzeCatalog = (products) => {
  const licenseRecommendations = [];
  const useCaseRecommendations = [];
  const priceIncreaseAlerts = [];

  products.forEach((product) => {
    const descLower = product.description.toLowerCase();

    // License modernization
    licenseModernizationRules.forEach((rule) => {
      if (rule.models.includes(product.licenseModel)) {
        licenseRecommendations.push({
          product: product.productFamily,
          ssPart: product.ssPart,
          licenseModel: product.licenseModel,
          title: rule.title,
          insight: rule.insight,
          consumptionRate: product.consumptionRate,
          price2029: product.price2029,
          price2030: product.price2030,
        });
      }
    });

    // Use case recommendations from market signals
    marketSignals.forEach((signal) => {
      const matched = signal.keywords.some((kw) => descLower.includes(kw));
      if (matched) {
        useCaseRecommendations.push({
          product: product.productFamily,
          ssPart: product.ssPart,
          title: signal.theme,
          demand: signal.demand,
          insight: `Align ${product.productFamily} messaging and packaging toward ${signal.theme.toLowerCase()} — a ${signal.demand.toLowerCase()}-demand priority in the current market.`,
          rationale: signal.summary,
          drivers: signal.drivers,
        });
      }
    });

    // Price increase alerts: 2030 price > 2029 price by more than 5%
    if (product.price2029 && product.price2030) {
      const pct = ((product.price2030 - product.price2029) / product.price2029) * 100;
      if (pct > 5) {
        priceIncreaseAlerts.push({
          product: product.productFamily,
          ssPart: product.ssPart,
          description: product.description,
          price2029: product.price2029,
          price2030: product.price2030,
          pctIncrease: pct,
        });
      }
    }
  });

  const categories = [...new Set(products.map((p) => p.category).filter(Boolean))];
  const licenseModelBreakdown = products.reduce((acc, p) => {
    acc[p.licenseModel] = (acc[p.licenseModel] ?? 0) + 1;
    return acc;
  }, {});

  return {
    licenseRecommendations,
    useCaseRecommendations,
    priceIncreaseAlerts,
    portfolioHighlights: {
      totalProducts: products.length,
      categories,
      licenseModelBreakdown,
      totalConsumptionRate: products.reduce((s, p) => s + (p.consumptionRate ?? 0), 0),
    },
  };
};

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------
function App() {
  const [products, setProducts] = useState(SAMPLE_ENTRIES);
  const [fileName, setFileName] = useState('Built-in sample (30 entries from ibm-license-catalog.xlsx)');
  const [parseError, setParseError] = useState(null);

  const analysis = useMemo(() => analyzeCatalog(products), [products]);

  const onFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setParseError(null);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const parsed = parseXlsxWorkbook(workbook);
      if (parsed.length === 0) {
        setParseError(
          'No IBM product rows were found. Make sure the file has a "Description" column with rows starting with "IBM".',
        );
        return;
      }
      setProducts(parsed);
      setFileName(`${file.name} (${parsed.length} entries)`);
    } catch (err) {
      setParseError(`Failed to parse file: ${err.message}`);
    }
  };

  const topLicenseModel = Object.entries(analysis.portfolioHighlights.licenseModelBreakdown).sort(
    (a, b) => b[1] - a[1],
  )[0]?.[0] ?? '—';

  return (
    <div className="app-shell">
      <header className="hero">
        <div>
          <p className="eyebrow">IBM Software Portfolio Advisor</p>
          <h1>Catalog modernization recommendations</h1>
          <p className="hero-copy">
            Upload <strong>ibm-license-catalog.xlsx</strong> (or any IBM ACP catalog XLSX) to get license
            modernization recommendations, market-aligned use case suggestions, and price change alerts — all
            derived directly from the real part numbers and pricing data.
          </p>
          <p className="file-badge">{fileName}</p>
        </div>
        <div className="hero-card">
          <h2>Upload catalog XLSX</h2>
          <p>Expects the IBM ACP format with columns: Description, S&amp;S part number, License part number, Royalty, Consumption rate, NTE prices for 2029 and 2030.</p>
          <label className="upload-button">
            Choose XLSX file
            <input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={onFileUpload} />
          </label>
          {parseError && <p className="parse-error">{parseError}</p>}
        </div>
      </header>

      <main className="layout">
        {/* Portfolio snapshot */}
        <section className="panel summary-panel full-width">
          <div className="section-heading">
            <h2>Portfolio snapshot</h2>
          </div>
          <div className="stats-grid">
            <article>
              <strong>{analysis.portfolioHighlights.totalProducts}</strong>
              <span>Line items</span>
            </article>
            <article>
              <strong>{analysis.licenseRecommendations.length}</strong>
              <span>License actions</span>
            </article>
            <article>
              <strong>{analysis.useCaseRecommendations.length}</strong>
              <span>Use-case ideas</span>
            </article>
            <article>
              <strong>{analysis.priceIncreaseAlerts.length}</strong>
              <span>Price alerts (&gt;5%)</span>
            </article>
            <article>
              <strong>{fmt(analysis.portfolioHighlights.totalConsumptionRate)}</strong>
              <span>Total consumption rate</span>
            </article>
            <article>
              <strong>{topLicenseModel}</strong>
              <span>Most common license</span>
            </article>
          </div>
          <div className="chips">
            {analysis.portfolioHighlights.categories.map((c) => (
              <span key={c} className="chip">{c}</span>
            ))}
          </div>
          <div className="chips">
            {Object.entries(analysis.portfolioHighlights.licenseModelBreakdown)
              .sort((a, b) => b[1] - a[1])
              .map(([model, count]) => (
                <span key={model} className="chip chip-secondary">{model} ({count})</span>
              ))}
          </div>
        </section>

        {/* License modernization */}
        <section className="panel recommendations-panel">
          <div className="section-heading">
            <h2>License modernization</h2>
            <span>{analysis.licenseRecommendations.length} recommendations</span>
          </div>
          {analysis.licenseRecommendations.length === 0 ? (
            <p className="empty-state">No license modernization patterns detected.</p>
          ) : (
            analysis.licenseRecommendations.map((item, i) => (
              <article className="recommendation-card" key={`lic-${i}`}>
                <div className="card-header">
                  <h3>{item.product}</h3>
                  <span className="badge">{item.licenseModel}</span>
                </div>
                <p className="recommendation-title">{item.title}</p>
                <p>{item.insight}</p>
                <div className="price-row">
                  <span>Consumption: <strong>{fmt(item.consumptionRate)}</strong></span>
                  <span>2029 NTE: <strong>{fmt(item.price2029)}</strong></span>
                  <span>2030 NTE: <strong>{fmt(item.price2030)}</strong></span>
                </div>
                <small>Part: {item.ssPart}</small>
              </article>
            ))
          )}
        </section>

        {/* Use case recommendations */}
        <section className="panel recommendations-panel">
          <div className="section-heading">
            <h2>New use cases from market signals</h2>
            <span>{analysis.useCaseRecommendations.length} ideas</span>
          </div>
          {analysis.useCaseRecommendations.length === 0 ? (
            <p className="empty-state">No market-aligned use cases matched.</p>
          ) : (
            analysis.useCaseRecommendations.map((item, i) => (
              <article className="recommendation-card" key={`uc-${i}`}>
                <div className="card-header">
                  <h3>{item.product}</h3>
                  <span className={`badge badge-demand-${item.demand.toLowerCase()}`}>{item.demand} demand</span>
                </div>
                <p className="recommendation-title">{item.title}</p>
                <p>{item.insight}</p>
                <small>{item.rationale}</small>
                <ul>
                  {item.drivers.map((d) => <li key={d}>{d}</li>)}
                </ul>
              </article>
            ))
          )}
        </section>

        {/* Price increase alerts */}
        <section className="panel recommendations-panel">
          <div className="section-heading">
            <h2>Price increase alerts (2029 → 2030)</h2>
            <span>{analysis.priceIncreaseAlerts.length} items</span>
          </div>
          {analysis.priceIncreaseAlerts.length === 0 ? (
            <p className="empty-state">No items with &gt;5% price increase detected.</p>
          ) : (
            analysis.priceIncreaseAlerts.map((item, i) => (
              <article className="recommendation-card" key={`pa-${i}`}>
                <div className="card-header">
                  <h3>{item.product}</h3>
                  <span className="badge badge-alert">+{item.pctIncrease.toFixed(1)}%</span>
                </div>
                <p className="recommendation-title">Proactive renewal or co-term opportunity</p>
                <p>Locking in the 2029 NTE price before renewal saves {fmt(item.price2030 - item.price2029)} per unit. Engage the client before the pricing period closes.</p>
                <div className="price-row">
                  <span>2029 NTE: <strong>{fmt(item.price2029)}</strong></span>
                  <span>2030 NTE: <strong>{fmt(item.price2030)}</strong></span>
                </div>
                <small>S&amp;S Part: {item.ssPart}</small>
              </article>
            ))
          )}
        </section>

        {/* Market trend watchlist */}
        <section className="panel trends-panel">
          <div className="section-heading">
            <h2>Market trend watchlist</h2>
          </div>
          <div className="trend-list">
            {marketSignals.map((signal) => (
              <article className="trend-card" key={signal.theme}>
                <div className="trend-header">
                  <h3>{signal.theme}</h3>
                  <span>{signal.demand} demand</span>
                </div>
                <p>{signal.summary}</p>
              </article>
            ))}
          </div>
        </section>
      </main>
    </div>
  );
}

export default App;
