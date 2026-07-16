import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
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
// License modernization patterns
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
// Helpers
// ---------------------------------------------------------------------------
const deriveProductFamily = (desc) => {
  if (!desc) return 'Unknown';
  const m = desc.match(
    /^IBM\s+([\w\s\-+]+?)\s+(?:Cartridge|Authorized|Floating|Virtual|Non-Production|Processor|Resource|Managed|Software|Developer|Simultaneous|Per\s|per\s|Suite|High|Data\s|Build\s|Deploy\s|Velocity|Committer)/i,
  );
  if (m) return `IBM ${m[1].trim()}`;
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
// XLSX parsing
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
  const sheetName =
    workbook.SheetNames.find((n) => /ACP/i.test(n)) ?? workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
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
// Built-in sample catalog
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

  const knowledgeBase = {
    totalSolutions: products.length,
    categories,
    topMarketSignals: marketSignals.slice(0, 3).map((s) => s.theme),
  };

  const executivePresentations = [...new Set(products.map((p) => p.productFamily))]
    .slice(0, 5)
    .map((family) => {
      const familyProducts = products.filter((p) => p.productFamily === family);
      const category = familyProducts[0]?.category ?? 'Other';
      const signal = marketSignals.find((s) => s.keywords.some((kw) => family.toLowerCase().includes(kw)));
      return {
        product: family,
        what: `${family} provides ${familyProducts.length} licensed capabilities for ${category.toLowerCase()}.`,
        why: signal
          ? signal.summary
          : 'Supports operational efficiency and digital transformation objectives.',
        how: `Deploy via existing entitlements. Integration with current environment requires architecture review and deployment planning with ATL/CSE support.`,
        whyNow: signal
          ? `Current market demand is ${signal.demand.toLowerCase()} for ${signal.theme.toLowerCase()}. Acting now aligns deployment with strategic priorities and renewal cycles.`
          : 'Current entitlement represents sunk cost with unrealized value. Immediate deployment maximizes ROI before renewal.',
      };
    });

  const deploymentPriorities = products
    .map((p) => ({
      product: p.productFamily,
      category: p.category,
      consumptionRate: p.consumptionRate ?? 0,
      price2029: p.price2029 ?? 0,
      licenseModel: p.licenseModel,
    }))
    .sort((a, b) => b.consumptionRate - a.consumptionRate)
    .slice(0, 10);

  const effortValueMatrix = useCaseRecommendations
    .slice(0, 8)
    .map((item) => {
      const priority = deploymentPriorities.find((entry) => entry.product === item.product);
      const valueScore = Math.min(100, Math.round(((priority?.consumptionRate ?? 0) / 150000) * 100) || (item.demand === 'High' ? 70 : 45));
      const effortBase = priority?.licenseModel === 'PVU / VPC' ? 75 : priority?.licenseModel === 'Authorized User' || priority?.licenseModel === 'Floating User' ? 45 : 60;
      const effortScore = Math.min(100, effortBase + (priority?.category === 'Security' || priority?.category === 'Integration' ? 10 : 0));
      const priorityBand =
        valueScore >= 65 && effortScore <= 50
          ? 'priority-1'
          : valueScore >= 65
            ? 'priority-2'
            : effortScore <= 50
              ? 'priority-3'
              : 'priority-4';
      return { ...item, valueScore, effortScore, priorityBand };
    });

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
    knowledgeBase,
    executivePresentations,
    deploymentPriorities,
    effortValueMatrix,
    products,
  };
};

// ---------------------------------------------------------------------------
// Tab definitions
// ---------------------------------------------------------------------------
const TABS = [
  { id: 'overview',       label: 'Overview' },
  { id: 'presentations',  label: 'Client Presentations' },
  { id: 'deployment',     label: 'Deployment Plan' },
  { id: 'plans',          label: 'Sales & Technical Plans' },
  { id: 'licensing',      label: 'Licensing & Cost' },
  { id: 'signals',        label: 'Market Signals' },
];

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// watsonx LLM endpoint config
// ---------------------------------------------------------------------------

// Relative path — proxied by Vite dev server (vite.config.js) in development
// and by nginx (openshift/nginx.conf) in production. Avoids browser CORS restrictions.
const WX_ENDPOINT =
  '/wx-api/ml/v1/deployments/wbplangenerator/text/generation_stream?version=2021-05-01';

function App() {
  const [products, setProducts] = useState(SAMPLE_ENTRIES);
  const [fileName, setFileName] = useState('Built-in sample (30 entries from ibm-license-catalog.xlsx)');
  const [parseError, setParseError] = useState(null);
  const [targetDate, setTargetDate] = useState('2025-12-31');
  const [selectedUseCaseIndex, setSelectedUseCaseIndex] = useState(0);
  const [activeTab, setActiveTab] = useState('overview');

  // ── watsonx workback plan LLM state ──────────────────────────────────────
  const [iamToken, setIamToken] = useState('');
  const [llmRows, setLlmRows] = useState(null);   // parsed JSON array once complete
  const [llmLoading, setLlmLoading] = useState(false);
  const [llmError, setLlmError] = useState(null);
  const abortRef = useRef(null);

  // Reset LLM output whenever the selected use case changes
  useEffect(() => {
    setLlmRows(null);
    setLlmError(null);
  }, [selectedUseCaseIndex]);

  const analysis = useMemo(() => analyzeCatalog(products), [products]);
  const selectedUseCase = analysis.useCaseRecommendations[selectedUseCaseIndex] ?? null;
  const selectedProductPriority =
    analysis.deploymentPriorities.find((item) => item.product === selectedUseCase?.product) ?? null;

  const salesPlan = useMemo(() => {
    if (!selectedUseCase) {
      return {
        priorityOpportunity: 'Select a use case in the Deployment Plan tab',
        renewalImpact: 'Select a use case to assess renewal and revenue impact.',
        executiveStrategy: 'Use the selected use case to shape the executive conversation.',
        expectedOutcome: 'No use case selected.',
      };
    }
    return {
      priorityOpportunity: `${selectedUseCase.product} — ${selectedUseCase.title}`,
      renewalImpact: `${selectedUseCase.product} supports ${selectedUseCase.title.toLowerCase()} and should be positioned before renewal events to improve realized value and reduce exposure to future price uplift.`,
      executiveStrategy: `Lead with ${selectedUseCase.title.toLowerCase()} as the board-level business priority, then map owned IBM entitlements to near-term deployment outcomes.`,
      expectedOutcome: `Increase deployment adoption of ${selectedUseCase.product} and convert owned entitlement into measurable business value.`,
    };
  }, [selectedUseCase]);

  const technicalPlan = useMemo(() => {
    if (!selectedUseCase) {
      return {
        deploymentSequence: ['Select a use case in the Deployment Plan tab to generate a technical plan'],
        prerequisites: 'No use case selected.',
        risks: 'No use case selected.',
        architecture: 'No use case selected.',
      };
    }
    return {
      deploymentSequence: [
        `1. Validate target architecture for ${selectedUseCase.product}`,
        `2. Confirm integration points required for ${selectedUseCase.title.toLowerCase()}`,
        `3. Launch pilot and success measurement for ${selectedUseCase.product}`,
      ],
      prerequisites: `Architecture review, entitlement validation, and environment readiness for ${selectedUseCase.product}.`,
      risks: `Integration complexity, deployment capacity, and adoption readiness for ${selectedUseCase.title.toLowerCase()}.`,
      architecture: selectedProductPriority
        ? `${selectedProductPriority.category} architecture with ${selectedProductPriority.licenseModel} deployment constraints and implementation planning.`
        : `Reference architecture should be defined for ${selectedUseCase.product}.`,
    };
  }, [selectedUseCase, selectedProductPriority]);

  // ── LLM streaming fetch ───────────────────────────────────────────────────
  const fetchLlmWorkbackPlan = useCallback(async () => {
    if (!selectedUseCase) return;
    if (!iamToken.trim()) {
      setLlmError('Please enter your IAM Bearer token before generating.');
      return;
    }

    // Abort any in-flight request
    if (abortRef.current) abortRef.current.abort();
    const controller = new AbortController();
    abortRef.current = controller;

    setLlmRows(null);
    setLlmError(null);
    setLlmLoading(true);

    const usecaseVar = `${selectedUseCase.product} — ${selectedUseCase.title}`;

    try {
      const response = await fetch(WX_ENDPOINT, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Authorization': `Bearer ${iamToken.trim()}`,
        },
        body: JSON.stringify({
          parameters: {
            prompt_variables: { Usecase: usecaseVar },
          },
        }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => response.statusText);
        throw new Error(`API error ${response.status}: ${text}`);
      }

      // Read the SSE / NDJSON stream — accumulate all tokens, parse JSON at the end
      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let lineBuffer = '';
      let fullText = '';

      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        lineBuffer += decoder.decode(value, { stream: true });

        // Each line may be an SSE "data: {...}" line or a raw JSON line
        const lines = lineBuffer.split('\n');
        lineBuffer = lines.pop(); // keep any incomplete trailing line

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed || trimmed === 'data: [DONE]') continue;
          const jsonStr = trimmed.startsWith('data:') ? trimmed.slice(5).trim() : trimmed;
          try {
            const chunk = JSON.parse(jsonStr);
            // watsonx generation_stream shape: { results: [{ generated_text: "..." }] }
            const token =
              chunk?.results?.[0]?.generated_text ??
              chunk?.generated_text ??
              '';
            fullText += token;
          } catch {
            // ignore non-JSON SSE envelope lines
          }
        }
      }

      // The model may echo the prompt before the real answer, producing multiple
      // JSON arrays in one output. Extract every [...] block and take the last one.
      const arrayMatches = [...fullText.matchAll(/\[[\s\S]*?\]/g)];
      if (arrayMatches.length === 0) {
        throw new Error(`Could not find a JSON array in the model response. Raw output:\n${fullText}`);
      }
      // Try each match from last to first — the final array is the actual answer
      let parsed = null;
      for (let i = arrayMatches.length - 1; i >= 0; i--) {
        try {
          const candidate = JSON.parse(arrayMatches[i][0]);
          if (Array.isArray(candidate) && candidate.length > 0 && candidate[0].week !== undefined) {
            parsed = candidate;
            break;
          }
        } catch {
          // not valid JSON, try the previous match
        }
      }
      if (!parsed) {
        throw new Error(`Could not parse the model response as a workback JSON array. Raw output:\n${fullText}`);
      }
      setLlmRows(parsed);
    } catch (err) {
      if (err.name !== 'AbortError') {
        setLlmError(err.message);
      }
    } finally {
      setLlmLoading(false);
    }
  }, [selectedUseCase, iamToken]);

  const topLicenseModel = Object.entries(analysis.portfolioHighlights.licenseModelBreakdown).sort(
    (a, b) => b[1] - a[1],
  )[0]?.[0] ?? '—';

  const onFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setParseError(null);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const parsed = parseXlsxWorkbook(workbook);
      if (parsed.length === 0) {
        setParseError('No IBM product rows were found. Make sure the file has a "Description" column with rows starting with "IBM".');
        return;
      }
      setProducts(parsed);
      setFileName(`${file.name} (${parsed.length} entries)`);
    } catch (err) {
      setParseError(`Failed to parse file: ${err.message}`);
    }
  };

  // ── Persistent header ────────────────────────────────────────────────────
  const renderHeader = () => (
    <header className="persistent-header">
      <div className="header-brand">
        <p className="eyebrow">IBM Software Portfolio Advisor — Powered by watsonx</p>
        <h1>Client-specific account planning &amp; deployment acceleration</h1>
      </div>
      <div className="header-controls">
        <div className="header-upload">
          <label className="upload-button">
            Upload catalog XLSX
            <input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={onFileUpload} />
          </label>
          <p className="file-badge">{fileName}</p>
          {parseError && <p className="parse-error">{parseError}</p>}
        </div>
        <div className="target-date-input">
          <label>Target deployment date</label>
          <input type="date" value={targetDate} onChange={(e) => setTargetDate(e.target.value)} />
        </div>
      </div>
    </header>
  );

  // ── Tab navigation ────────────────────────────────────────────────────────
  const renderTabs = () => (
    <nav className="tab-nav" role="tablist">
      {TABS.map((tab) => (
        <button
          key={tab.id}
          role="tab"
          aria-selected={activeTab === tab.id}
          className={`tab-btn${activeTab === tab.id ? ' tab-btn-active' : ''}`}
          onClick={() => setActiveTab(tab.id)}
          type="button"
        >
          {tab.label}
        </button>
      ))}
    </nav>
  );

  // ── Tab 1: Overview ───────────────────────────────────────────────────────
  const renderOverview = () => (
    <div className="tab-content">
      {/* Phase 1: Knowledge base */}
      <section className="panel phase-section">
        <div className="section-heading">
          <h2>Phase 1: Client-specific knowledge base</h2>
        </div>
        <p>Total solutions in catalog: <strong>{analysis.knowledgeBase.totalSolutions}</strong></p>
        <p>Solution categories: <strong>{analysis.knowledgeBase.categories.join(', ')}</strong></p>
        <p>Top market signals: <strong>{analysis.knowledgeBase.topMarketSignals.join(', ')}</strong></p>
      </section>

      {/* Portfolio snapshot */}
      <section className="panel summary-panel">
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
    </div>
  );

  // ── Tab 2: Client Presentations ───────────────────────────────────────────
  const renderPresentations = () => (
    <div className="tab-content">
      <section className="panel recommendations-panel">
        <div className="section-heading">
          <h2>Executive client presentations (What / Why / How / Why Now)</h2>
          <span>{analysis.executivePresentations.length} products</span>
        </div>
        {analysis.executivePresentations.map((item, i) => (
          <article className="recommendation-card" key={`exec-${i}`}>
            <h3 className="recommendation-title">{item.product}</h3>
            <p><strong>What:</strong> {item.what}</p>
            <p><strong>Why:</strong> {item.why}</p>
            <p><strong>How:</strong> {item.how}</p>
            <p><strong>Why Now:</strong> {item.whyNow}</p>
          </article>
        ))}
      </section>
    </div>
  );

  // ── Tab 3: Deployment Plan ────────────────────────────────────────────────
  const renderDeploymentPlan = () => (
    <div className="tab-content">
      {/* Phase 3: Workback plan */}
      <section className="panel phase-section">
        <div className="section-heading">
          <h2>Phase 3: Workback plan to target deployment ({targetDate})</h2>
        </div>
        <div className="section-heading">
          <h3>Selected use case</h3>
          <span>{selectedUseCase ? `${selectedUseCase.product} — ${selectedUseCase.title}` : 'None selected'}</span>
        </div>
        <div className="use-case-selector">
          {analysis.useCaseRecommendations.slice(0, 5).map((item, index) => (
            <button
              type="button"
              key={`selector-${item.product}-${index}`}
              className={`use-case-option${selectedUseCaseIndex === index ? ' use-case-option-active' : ''}`}
              onClick={() => setSelectedUseCaseIndex(index)}
            >
              <strong>{item.product}</strong>
              <span>{item.title}</span>
            </button>
          ))}
        </div>

        {/* ── watsonx LLM workback plan ─────────────────────────────────── */}
        <div className="llm-workback-panel">
          <div className="llm-workback-header">
            <span className="llm-badge">Powered by watsonx</span>
            <p className="llm-description">
              Generate an AI workback plan for <strong>{selectedUseCase ? `${selectedUseCase.product} — ${selectedUseCase.title}` : 'the selected use case'}</strong> using your deployed watsonx LLM.
            </p>
          </div>
          <div className="llm-controls">
            <input
              type="password"
              className="llm-token-input"
              placeholder="Paste your IAM Bearer token"
              value={iamToken}
              onChange={(e) => setIamToken(e.target.value)}
              autoComplete="off"
              spellCheck={false}
            />
            <button
              type="button"
              className={`llm-generate-btn${llmLoading ? ' llm-generate-btn-loading' : ''}`}
              onClick={fetchLlmWorkbackPlan}
              disabled={llmLoading || !selectedUseCase}
            >
              {llmLoading ? 'Generating…' : 'Generate with watsonx'}
            </button>
            {llmLoading && (
              <button
                type="button"
                className="llm-stop-btn"
                onClick={() => { abortRef.current?.abort(); setLlmLoading(false); }}
              >
                Stop
              </button>
            )}
          </div>
          {llmError && <p className="llm-error">{llmError}</p>}
          {llmLoading && (
            <div className="llm-streaming-indicator">
              <span className="llm-cursor">▋</span> Generating workback plan…
            </div>
          )}
          {llmRows && !llmLoading && (
            <div className="table-scroll llm-table-output">
              <table className="priority-table">
                <thead>
                  <tr>
                    <th>Week</th>
                    <th>Activity</th>
                    <th>Owner</th>
                    <th>Target Date</th>
                    <th>Focus</th>
                    <th>Dependencies</th>
                  </tr>
                </thead>
                <tbody>
                  {llmRows.map((row, i) => (
                    <tr key={`llm-row-${i}`}>
                      <td>W{row.week}</td>
                      <td>{row.activity}</td>
                      <td>{row.owner}</td>
                      <td>{row.targetDate || '—'}</td>
                      <td>{row.focus}</td>
                      <td>{row.dependencies}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

      </section>

      {/* Phase 4: Value prioritization */}
      <section className="panel phase-section">
        <div className="section-heading">
          <h2>Phase 4: Deployment value prioritization</h2>
          <span>Top {analysis.deploymentPriorities.length} by consumption rate</span>
        </div>
        <div className="matrix-legend">
          <span className="matrix-chip matrix-chip-priority-1">Priority 1: high value / low effort</span>
          <span className="matrix-chip matrix-chip-priority-2">Priority 2: high value / higher effort</span>
          <span className="matrix-chip matrix-chip-priority-3">Priority 3: moderate value / low effort</span>
          <span className="matrix-chip matrix-chip-priority-4">Priority 4: lower priority</span>
        </div>
        <div className="matrix-grid">
          {analysis.effortValueMatrix.map((item, index) => (
            <article
              key={`matrix-${item.product}-${index}`}
              className={`matrix-card ${item.priorityBand}${selectedUseCase?.product === item.product && selectedUseCase?.title === item.title ? ' matrix-card-selected' : ''}`}
            >
              <h3>{item.product}</h3>
              <p>{item.title}</p>
              <p><strong>Value:</strong> {item.valueScore}/100</p>
              <p><strong>Effort:</strong> {item.effortScore}/100</p>
            </article>
          ))}
        </div>
      </section>
    </div>
  );

  // ── Tab 4: Sales & Technical Plans ───────────────────────────────────────
  const renderPlans = () => (
    <div className="tab-content tab-content-split">
      <section className="panel recommendations-panel phase-section">
        <div className="section-heading">
          <h2>Phase 5a: Business plan for technology sales leaders</h2>
        </div>
        <p><strong>Priority opportunity:</strong> {salesPlan.priorityOpportunity}</p>
        <p><strong>Expected outcome:</strong> {salesPlan.expectedOutcome}</p>
        <p><strong>Renewal impact:</strong> {salesPlan.renewalImpact}</p>
        <p><strong>Executive strategy:</strong> {salesPlan.executiveStrategy}</p>
      </section>
      <section className="panel recommendations-panel phase-section">
        <div className="section-heading">
          <h2>Phase 5b: Technical plan for ATLs and CSEs</h2>
        </div>
        <p><strong>Recommended deployment sequence:</strong></p>
        <ul>
          {technicalPlan.deploymentSequence.map((s, i) => <li key={i}>{s}</li>)}
        </ul>
        <p><strong>Architecture considerations:</strong> {technicalPlan.architecture}</p>
        <p><strong>Prerequisites:</strong> {technicalPlan.prerequisites}</p>
        <p><strong>Risks:</strong> {technicalPlan.risks}</p>
      </section>
    </div>
  );

  // ── Tab 5: Licensing & Cost ───────────────────────────────────────────────
  const renderLicensing = () => (
    <div className="tab-content">
      {/* Product catalog table */}
      <section className="panel">
        <div className="section-heading">
          <h2>Product catalog</h2>
          <span>{analysis.products.length} line items</span>
        </div>
        <div className="table-scroll">
          <table className="priority-table">
            <thead>
              <tr>
                <th>Product</th>
                <th>Category</th>
                <th>License Model</th>
                <th>Consumption Rate</th>
                <th>2029 NTE</th>
              </tr>
            </thead>
            <tbody>
              {analysis.deploymentPriorities.map((item, i) => (
                <tr key={`cat-${i}`}>
                  <td>{item.product}</td>
                  <td>{item.category}</td>
                  <td>{item.licenseModel}</td>
                  <td>{fmt(item.consumptionRate)}</td>
                  <td>{fmt(item.price2029)}</td>
                </tr>
              ))}
            </tbody>
          </table>
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
    </div>
  );

  // ── Tab 6: Market Signals ─────────────────────────────────────────────────
  const renderSignals = () => (
    <div className="tab-content tab-content-split">
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
    </div>
  );

  const tabRenderers = {
    overview: renderOverview,
    presentations: renderPresentations,
    deployment: renderDeploymentPlan,
    plans: renderPlans,
    licensing: renderLicensing,
    signals: renderSignals,
  };

  return (
    <div className="app-shell">
      {renderHeader()}
      {renderTabs()}
      <main>
        {tabRenderers[activeTab]?.()}
      </main>
    </div>
  );
}

export default App;
