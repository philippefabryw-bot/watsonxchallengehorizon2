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
// XLSX parsing — reference catalog format (Weighting sheet)
// Header row contains: S&S Part No, Licence Part No, Licence Description,
// Deployment Plan Quantity, z (platform), License Entitled Price, ...
// Drawdown Price - Discounted = IBM cost
// Drawdown Price - Weighted   = client price
// Deployment Plan Quantity non-null/non-zero = deployed; null/0 = available
// ---------------------------------------------------------------------------
const parseXlsxWorkbook = (workbook) => {
  // Prefer the "Weighting" sheet; fall back to first sheet
  const sheetName =
    workbook.SheetNames.find((n) => /weighting/i.test(n) && !/BL/i.test(n)) ??
    workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

  // Find the header row: must contain both "Licence Description" and "Deployment Plan Quantity"
  const headerRowIndex = raw.findIndex((row) => {
    const joined = row.map((v) => String(v ?? '')).join('|');
    return joined.includes('Licence Description') && joined.includes('Deployment Plan Quantity');
  });
  if (headerRowIndex < 0) return [];

  const headers = raw[headerRowIndex].map((h) => String(h ?? '').trim());
  const hIdx = (name) => headers.findIndex((h) => h === name);
  const hMatch = (pat) => headers.findIndex((h) => pat.test(h));

  const ssPartIdx        = hIdx('S&S Part No');
  const licPartIdx       = hIdx('Licence Part No');
  const descIdx          = hIdx('Licence Description');
  const deployQtyIdx     = hIdx('Deployment Plan Quantity');
  const platformIdx      = hIdx('z');
  const licEntitledIdx   = hIdx('License Entitled Price');
  const ssEntitledIdx    = hIdx('S&S Entitled Price');
  const drawEntitledIdx  = hIdx('Drawdown Price - Entitled');
  const discountIdx      = hIdx('Discount');
  const ibmCostIdx       = hIdx('Drawdown Price - Discounted');
  const clientPriceIdx   = hIdx('Drawdown Price - Weighted');
  const extNetIdx        = hIdx('Extended Net Price');
  const extWeightedIdx   = hIdx('Extended Weighted Price');

  const safeNum = (row, i) => {
    if (i < 0 || i >= row.length) return null;
    const v = row[i];
    if (v === null || v === undefined) return null;
    if (typeof v === 'string' && (v.startsWith('#') || !v.trim())) return null;
    const n = Number(v);
    return isNaN(n) ? null : n;
  };
  const safeStr = (row, i) => {
    if (i < 0 || i >= row.length) return null;
    const v = row[i];
    if (v === null || v === undefined) return null;
    if (typeof v === 'string' && v.startsWith('#')) return null;
    return String(v).trim() || null;
  };

  return raw
    .slice(headerRowIndex + 1)
    .filter((row) => {
      const desc = safeStr(row, descIdx);
      return desc && (safeStr(row, licPartIdx) || safeStr(row, ssPartIdx));
    })
    .map((row, index) => {
      const desc = safeStr(row, descIdx);
      const deployQty = safeNum(row, deployQtyIdx);
      return {
        id: `row-${index}`,
        ssPart:           safeStr(row, ssPartIdx) ?? '',
        licensePart:      safeStr(row, licPartIdx) ?? '',
        description:      desc,
        platform:         safeStr(row, platformIdx) ?? 'PA',
        deploymentQty:    deployQty,
        isDeployed:       deployQty !== null && deployQty > 0,
        licenceEntitled:  safeNum(row, licEntitledIdx),
        ssEntitled:       safeNum(row, ssEntitledIdx),
        drawdownEntitled: safeNum(row, drawEntitledIdx),
        discount:         safeNum(row, discountIdx),
        ibmCost:          safeNum(row, ibmCostIdx),
        clientPrice:      safeNum(row, clientPriceIdx),
        extendedNetPrice: safeNum(row, extNetIdx),
        extendedWeighted: safeNum(row, extWeightedIdx),
        productFamily:    deriveProductFamily(desc),
        licenseModel:     deriveLicenseModel(desc),
        category:         deriveCategory(desc),
      };
    });
};

// ---------------------------------------------------------------------------
// Built-in sample catalog — sourced from reference catalog.xlsx (Weighting sheet)
// Deployed: Deployment Plan Quantity > 0  |  Available: Deployment Plan Quantity = 0
// ibmCost   = Drawdown Price - Discounted (what IBM charges)
// clientPrice = Drawdown Price - Weighted (what the client pays)
// ---------------------------------------------------------------------------
const SAMPLE_ENTRIES = [
  // ── Deployed (Deployment Plan Quantity > 0) ─────────────────────────────
  { ssPart: 'E11R5ZX', licPart: 'D11R6ZX', desc: 'IBM webMethods Hybrid Integration for IBM Z Resource Unit License + SW Subscription & Support 12 Months',                                        deployQty: 1436, platform: 'PA', ibmCost: 213.94,    clientPrice: 1404.44,     extNet: 307217.84,    extWeighted: 2016770.26 },
  { ssPart: 'E0ZP5ZX', licPart: 'D0ZP7ZX', desc: 'IBM Enterprise Application Runtimes for IBM Z Virtual Processor Core License + Software Subscription & Support 12 Months',                     deployQty: 39,   platform: 'PA', ibmCost: 5669.92,   clientPrice: 37220.88,    extNet: 221126.75,    extWeighted: 1451614.45 },
  { ssPart: 'E0LM2LL', licPart: 'D1GV2LL', desc: 'IBM Verify Identity Access Virtual Edition per Processor Value Unit for IBM Z License + SW Subscription & Support 12 Months',                  deployQty: 592,  platform: 'PA', ibmCost: 346.75,    clientPrice: 2276.30,     extNet: 205277.01,    extWeighted: 1347566.80 },
  { ssPart: 'E0REPZX', licPart: 'D0RERZX', desc: 'IBM watsonx Orchestrate On-Prem for IBM Z Install LIC + SW S&S 12 Months',                                                                     deployQty: 1,    platform: 'PA', ibmCost: 164925.73, clientPrice: 1082675.73,  extNet: 164925.73,    extWeighted: 1082675.73 },
  { ssPart: 'E0RGNZX', licPart: 'D0RGQZX', desc: 'IBM watsonx Orchestrate On-Prem for IBM Z RU 1000 Resource Unit LIC + SW S&S 12 Months',                                                       deployQty: 1,    platform: 'PA', ibmCost: 23672.63,  clientPrice: 155401.96,   extNet: 23672.63,     extWeighted: 155401.96 },
  { ssPart: 'E0HAFZX', licPart: 'D0HANZX', desc: 'IBM watsonx.governance Operational Risk Management for IBM Z Application License + Software Subscription & Support 12Months',                   deployQty: 2,    platform: 'PA', ibmCost: 54778.66,  clientPrice: 359601.43,   extNet: 109557.32,    extWeighted: 719202.86 },
  { ssPart: 'E0HAMZX', licPart: 'D0HAGZX', desc: 'IBM watsonx.governance Model Risk Governance for IBM Z Application License + Software Subscription & Support 12Months',                         deployQty: 2,    platform: 'PA', ibmCost: 47070.68,  clientPrice: 309001.45,   extNet: 94141.37,     extWeighted: 618002.89 },
  { ssPart: 'E0HAUZX', licPart: 'D0HAVZX', desc: 'IBM watsonx.governance Regulatory Compliance Management for IBM Z Application License + Software Subscription & Support 12Months',              deployQty: 1,    platform: 'PA', ibmCost: 62740.44,  clientPrice: 411867.52,   extNet: 62740.44,     extWeighted: 411867.52 },
  { ssPart: 'E14STZX', licPart: 'D14SVZX', desc: 'IBM Db2 AI Advanced Edition for Linux on IBM Z Virtual Processor Core LIC + SW S&S 12 MO',                                                     deployQty: 5,    platform: 'PA', ibmCost: 17740.64,  clientPrice: 116460.65,   extNet: 88703.18,     extWeighted: 582303.24 },
  { ssPart: 'E10KIZX', licPart: 'D10KKZX', desc: 'IBM watsonx.data intelligence for IBM Z Resource Unit License + SW Subscription & Support 12 Months',                                           deployQty: 956,  platform: 'PA', ibmCost: 108.88,    clientPrice: 714.78,      extNet: 104093.30,    extWeighted: 683333.57 },
  { ssPart: 'E10E1ZX', licPart: 'D10E3ZX', desc: 'IBM watsonx.data integration for IBM Z Resource Unit License + SW Subscription & Support 12 Months',                                            deployQty: 175,  platform: 'PA', ibmCost: 626.96,    clientPrice: 4115.74,     extNet: 109717.44,    extWeighted: 720253.97 },
  { ssPart: '5655E26',  licPart: '5655E26',  desc: 'IMS HP Unload',                                                                                                                               deployQty: 22,   platform: 'z',  ibmCost: 226.91,    clientPrice: 1489.58,     extNet: 4992.03,      extWeighted: 32770.80 },
  { ssPart: '5655E29',  licPart: '5655U09',  desc: 'IMS HP Pointer Checker V3',                                                                                                                   deployQty: 22,   platform: 'z',  ibmCost: 245.89,    clientPrice: 1614.17,     extNet: 5409.56,      extWeighted: 35511.74 },
  { ssPart: '5655E90',  licPart: '5698AT3',  desc: 'Db2 Administration Tool 13',                                                                                                                  deployQty: 22,   platform: 'z',  ibmCost: 988.25,    clientPrice: 6487.46,     extNet: 21741.41,     extWeighted: 142724.20 },
  // ── Available in catalog (Deployment Plan Quantity = 0) ─────────────────
  { ssPart: 'E0KHLLL', licPart: 'D1B0JLL', desc: 'Application Optimization Module for IBM DataPower Gateway Virtual Edition Processor Value Unit License + Software Subscription & Support 12 Months', deployQty: 0, platform: 'PA', ibmCost: 42.79,     clientPrice: 280.89,      extNet: 0, extWeighted: 0 },
  { ssPart: 'E0PMHLL', licPart: 'D20RULL', desc: 'IBM Cloud Pak for Business Automation Virtual Processor Core License + SW Subscription & Support 12 Months',                                    deployQty: 0,    platform: 'PA', ibmCost: 10461.89,  clientPrice: 68678.40,    extNet: 0, extWeighted: 0 },
  { ssPart: 'E0QXFLL', licPart: 'D27RMLL', desc: 'IBM Cloud Pak for Data Enterprise Edition Non-Production Virtual Processor Core License + SW Subscriptn and Support 12 Months',                  deployQty: 0,    platform: 'PA', ibmCost: 4331.55,   clientPrice: 28435.03,    extNet: 0, extWeighted: 0 },
  { ssPart: 'E0PD4LL', licPart: 'D1YGZLL', desc: 'IBM Cloud Pak for Data Enterprise Edition per Virtual Processor Core License + SW Subscription & Support 12 Months',                             deployQty: 0,    platform: 'PA', ibmCost: 8663.22,   clientPrice: 56870.79,    extNet: 0, extWeighted: 0 },
  { ssPart: 'E0PPCLL', licPart: 'D20ZILL', desc: 'IBM Cloud Pak for Integration Virtual Processor Core for IBM Z License + SW Subscription & Support 12 Months',                                   deployQty: 0,    platform: 'PA', ibmCost: 13618.86,  clientPrice: 89402.71,    extNet: 0, extWeighted: 0 },
  { ssPart: 'E0PPBLL', licPart: 'D20ZBLL', desc: 'IBM Cloud Pak for Integration Virtual Processor Core License + SW Subscription & Support 12 Months',                                             deployQty: 0,    platform: 'PA', ibmCost: 11088.82,  clientPrice: 72820.39,    extNet: 0, extWeighted: 0 },
  { ssPart: 'E0REPZX', licPart: 'D0REZOX', desc: 'IBM watsonx Orchestrate On-Prem for IBM Z Small Install LIC + SW S&S 12 Months',                                                                deployQty: 0,    platform: 'PA', ibmCost: 55183.68,  clientPrice: 362223.57,   extNet: 0, extWeighted: 0 },
  { ssPart: 'E0HAUZX', licPart: 'D0HAWZX', desc: 'IBM watsonx.governance Financial Crimes Compliance for IBM Z Application License + Software Subscription & Support 12Months',                   deployQty: 0,    platform: 'PA', ibmCost: 54778.66,  clientPrice: 359601.43,   extNet: 0, extWeighted: 0 },
  { ssPart: 'E14STZX', licPart: 'D14SUZZ', desc: 'IBM Db2 AI Advanced Edition for IBM Z Virtual Processor Core LIC + SW S&S 12 MO',                                                              deployQty: 0,    platform: 'z',  ibmCost: 17140.64,  clientPrice: 112560.65,   extNet: 0, extWeighted: 0 },
  { ssPart: 'E10KIZX', licPart: 'D10KLZX', desc: 'IBM watsonx.data intelligence for IBM Z 10K Resource Unit License + SW Subscription & Support 12 Months',                                       deployQty: 0,    platform: 'PA', ibmCost: 1088.80,   clientPrice: 7147.84,     extNet: 0, extWeighted: 0 },
  { ssPart: 'E0ZNZZX', licPart: 'D0ZP1ZX', desc: 'IBM Enterprise Application Runtimes Virtual Processor Core License + Software Subscription & Support 12 Months',                                deployQty: 0,    platform: 'PA', ibmCost: 3718.39,   clientPrice: 24408.58,    extNet: 0, extWeighted: 0 },
  { ssPart: 'E034QZX', licPart: 'D034RZX', desc: 'IBM Information Server Cartridge for Virtual Processor Core License + SW Subscription & Support 12 Months',                                     deployQty: 0,    platform: 'PA', ibmCost: 20637.50,  clientPrice: 135484.38,   extNet: 0, extWeighted: 0 },
  { ssPart: 'E03JRZX', licPart: 'D03JSZX', desc: 'IBM DataStage Enterprise Plus Cartridge for Virtual Processor Core License + SW Subscription and Support 12 Months',                            deployQty: 0,    platform: 'PA', ibmCost: 12890.49,  clientPrice: 84614.42,    extNet: 0, extWeighted: 0 },
  { ssPart: 'E0AE3ZX', licPart: 'D0AE4ZX', desc: 'IBM Security QRadar Suite Software 100 Resource Unit License + Subscription & Support 12 Months',                                               deployQty: 0,    platform: 'PA', ibmCost: 675.33,    clientPrice: 4432.16,     extNet: 0, extWeighted: 0 },
  { ssPart: 'E09PILL', licPart: 'D0EKZLL', desc: 'IBM SPSS Statistics Standard Authorized User License + SW Subscription & Support 12 Months',                                                    deployQty: 0,    platform: 'PA', ibmCost: 1393.58,   clientPrice: 9149.46,     extNet: 0, extWeighted: 0 },
  { ssPart: 'E0N6PLL', licPart: 'D1QZVLL', desc: 'IBM ILOG CPLEX Optimizer Cluster Edition per Process License + SW Subscription & Support 12 Months',                                            deployQty: 0,    platform: 'PA', ibmCost: 1557.12,   clientPrice: 10220.89,    extNet: 0, extWeighted: 0 },
].map((e, i) => ({
  id: `sample-${i}`,
  ssPart: e.ssPart,
  licensePart: e.licPart,
  description: e.desc,
  platform: e.platform,
  deploymentQty: e.deployQty,
  isDeployed: e.deployQty > 0,
  ibmCost: e.ibmCost,
  clientPrice: e.clientPrice,
  extendedNetPrice: e.extNet,
  extendedWeighted: e.extWeighted,
  productFamily: deriveProductFamily(e.desc),
  licenseModel: deriveLicenseModel(e.desc),
  category: deriveCategory(e.desc),
}));

// ---------------------------------------------------------------------------
// Analysis engine
// ---------------------------------------------------------------------------
const fmt = (n) =>
  n == null ? '—' : `$${Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

const fmtQty = (n) => (n == null ? '—' : Number(n).toLocaleString('en-US'));

const analyzeCatalog = (products) => {
  const licenseRecommendations = [];
  const useCaseRecommendations = [];

  // Split into deployed vs available
  const deployedProducts   = products.filter((p) => p.isDeployed);
  const availableProducts  = products.filter((p) => !p.isDeployed);

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
          ibmCost: product.ibmCost,
          clientPrice: product.clientPrice,
          isDeployed: product.isDeployed,
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

  // Deployment priorities ordered by client price (highest value first)
  const deploymentPriorities = products
    .map((p) => ({
      product: p.productFamily,
      description: p.description,
      category: p.category,
      platform: p.platform,
      deploymentQty: p.deploymentQty,
      isDeployed: p.isDeployed,
      ibmCost: p.ibmCost ?? 0,
      clientPrice: p.clientPrice ?? 0,
      licenseModel: p.licenseModel,
      ssPart: p.ssPart,
      licensePart: p.licensePart,
    }))
    .sort((a, b) => b.clientPrice - a.clientPrice)
    .slice(0, 20);

  const effortValueMatrix = useCaseRecommendations
    .slice(0, 8)
    .map((item) => {
      const priority = deploymentPriorities.find((entry) => entry.product === item.product);
      const valueScore = Math.min(100, Math.round(((priority?.clientPrice ?? 0) / 1200000) * 100) || (item.demand === 'High' ? 70 : 45));
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

  // Totals
  const totalExtendedNet      = deployedProducts.reduce((s, p) => s + (p.extendedNetPrice ?? 0), 0);
  const totalExtendedWeighted = deployedProducts.reduce((s, p) => s + (p.extendedWeighted ?? 0), 0);

  return {
    licenseRecommendations,
    useCaseRecommendations,
    portfolioHighlights: {
      totalProducts: products.length,
      deployedCount: deployedProducts.length,
      availableCount: availableProducts.length,
      categories,
      licenseModelBreakdown,
      totalExtendedNet,
      totalExtendedWeighted,
    },
    knowledgeBase,
    executivePresentations,
    deploymentPriorities,
    effortValueMatrix,
    deployedProducts,
    availableProducts,
    products,
  };
};

// ---------------------------------------------------------------------------
// Tab definitions
// ---------------------------------------------------------------------------
const TABS = [
  { id: 'overview',       label: 'Overview' },
  { id: 'deployed',       label: 'Deployed' },
  { id: 'available',      label: 'Available Catalog' },
  { id: 'presentations',  label: 'Client Presentations' },
  { id: 'deployment',     label: 'Deployment Plan' },
  { id: 'plans',          label: 'Sales & Technical Plans' },
  { id: 'licensing',      label: 'Licensing & Cost' },
  { id: 'signals',        label: 'Market Signals' },
];

// ---------------------------------------------------------------------------
// Markdown table → llmRows parser
// Handles the rich markdown table the model returns when it doesn't emit JSON.
// Maps any column names the model uses to the fixed { week, activity, owner,
// targetDate, focus, dependencies } shape expected by the table renderer.
// ---------------------------------------------------------------------------
function parseMarkdownTable(text) {
  // Find the first line that looks like a markdown table header (contains |)
  const lines = text.split('\n').map((l) => l.trim()).filter(Boolean);
  const headerIdx = lines.findIndex((l) => l.startsWith('|') && l.includes('|', 1));
  if (headerIdx === -1) return null;

  // Parse column names from the header row
  const headers = lines[headerIdx]
    .split('|')
    .map((h) => h.trim().toLowerCase().replace(/[^a-z0-9 ]/g, ''))
    .filter(Boolean);

  // Skip the separator row (---|---) and collect data rows
  const dataLines = lines.slice(headerIdx + 1).filter(
    (l) => l.startsWith('|') && !/^\|[\s|:-]+\|$/.test(l),
  );

  if (dataLines.length === 0) return null;

  // Column index helpers — pick the best match for each target field
  const col = (candidates) => {
    for (const c of candidates) {
      const idx = headers.findIndex((h) => h.includes(c));
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const iPhase      = col(['phase', 'week', 'step']);
  const iActivity   = col(['key activit', 'activit', 'task', 'description']);
  const iOwner      = col(['owner']);
  const iStart      = col(['start', 'relative', 'target date', 'date']);
  const iDuration   = col(['duration', 'effort', 'focus']);
  const iMilestone  = col(['milestone', 'deliverable', 'depend', 'outcome']);

  return dataLines.map((line, i) => {
    const cells = line
      .split('|')
      .map((c) => c.trim().replace(/<br\s*\/?>/gi, ' ').replace(/\*\*/g, '').replace(/[🔟1️⃣2️⃣3️⃣4️⃣5️⃣6️⃣7️⃣8️⃣9️⃣]/g, '').trim())
      .filter((_, idx) => idx > 0 && idx <= headers.length); // drop leading/trailing empty splits

    const get = (idx) => (idx !== -1 && idx < cells.length ? cells[idx] || '—' : '—');

    return {
      week:         get(iPhase)     || String(i + 1),
      activity:     get(iActivity),
      owner:        get(iOwner),
      targetDate:   get(iStart),
      focus:        get(iDuration),
      dependencies: get(iMilestone),
    };
  });
}

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------
function App() {
  const [products, setProducts] = useState(SAMPLE_ENTRIES);
  const [fileName, setFileName] = useState('Built-in sample (29 entries from reference catalog.xlsx — 14 deployed, 15 available)');
  const [parseError, setParseError] = useState(null);
  const [targetDate, setTargetDate] = useState('2025-12-31');
  const [selectedUseCaseIndex, setSelectedUseCaseIndex] = useState(0);
  const [activeTab, setActiveTab] = useState('overview');

  // ── watsonx workback plan LLM state ──────────────────────────────────────
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

  // ── LLM fetch (via backend proxy) ────────────────────────────────────────
  const fetchLlmWorkbackPlan = useCallback(async () => {
    if (!selectedUseCase) return;

    if (abortRef.current) abortRef.current.abort();
    const controller = new AbortController();
    abortRef.current = controller;

    setLlmRows(null);
    setLlmError(null);
    setLlmLoading(true);

    const usecaseVar = `${selectedUseCase.product} — ${selectedUseCase.title}`;

    try {
      const response = await fetch('/api/workback', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ usecase: usecaseVar }),
        signal: controller.signal,
      });

      if (!response.ok) {
        const text = await response.text().catch(() => response.statusText);
        throw new Error(`API error ${response.status}: ${text}`);
      }

      const data = await response.json();
      // watsonx generation (non-stream) shape: { results: [{ generated_text }] }
      const fullText =
        data?.results?.[0]?.generated_text ??
        data?.generated_text ??
        JSON.stringify(data, null, 2);

      // ── 1. Try JSON array (original format) ────────────────────────────
      let parsed = null;
      const arrayMatches = [...fullText.matchAll(/\[[\s\S]*?\]/g)];
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

      // ── 2. Fall back: parse a markdown table ────────────────────────────
      if (!parsed) {
        parsed = parseMarkdownTable(fullText);
      }

      if (!parsed || parsed.length === 0) {
        throw new Error(`Could not parse the model response. Raw output:\n${fullText}`);
      }
      setLlmRows(parsed);
    } catch (err) {
      if (err.name !== 'AbortError') {
        setLlmError(err.message);
      }
    } finally {
      setLlmLoading(false);
    }
  }, [selectedUseCase]);

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
            <strong>{analysis.portfolioHighlights.deployedCount}</strong>
            <span>Deployed</span>
          </article>
          <article>
            <strong>{analysis.portfolioHighlights.availableCount}</strong>
            <span>Available (not deployed)</span>
          </article>
          <article>
            <strong>{analysis.licenseRecommendations.length}</strong>
            <span>License actions</span>
          </article>
          <article>
            <strong>{fmt(analysis.portfolioHighlights.totalExtendedNet)}</strong>
            <span>Total IBM cost (deployed)</span>
          </article>
          <article>
            <strong>{fmt(analysis.portfolioHighlights.totalExtendedWeighted)}</strong>
            <span>Total client value (deployed)</span>
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

  // ── Tab 2: Deployed ───────────────────────────────────────────────────────
  const renderDeployed = () => (
    <div className="tab-content">
      <section className="panel">
        <div className="section-heading">
          <h2>Deployed products</h2>
          <span>{analysis.deployedProducts.length} items with active deployment</span>
        </div>
        <p className="panel-note">Products where <strong>Deployment Plan Quantity &gt; 0</strong>. IBM cost = Drawdown Price – Discounted. Client price = Drawdown Price – Weighted.</p>
        <div className="table-scroll">
          <table className="priority-table">
            <thead>
              <tr>
                <th>Licence Part</th>
                <th>Description</th>
                <th>Platform</th>
                <th>Deploy Qty</th>
                <th>IBM Cost (unit)</th>
                <th>Client Price (unit)</th>
                <th>Extended IBM Cost</th>
                <th>Extended Client Value</th>
                <th>License Model</th>
              </tr>
            </thead>
            <tbody>
              {analysis.deployedProducts.map((p, i) => (
                <tr key={`dep-${i}`}>
                  <td><code>{p.licensePart}</code></td>
                  <td>{p.description}</td>
                  <td><span className={`badge badge-platform-${(p.platform||'').toLowerCase()}`}>{p.platform}</span></td>
                  <td>{fmtQty(p.deploymentQty)}</td>
                  <td>{fmt(p.ibmCost)}</td>
                  <td>{fmt(p.clientPrice)}</td>
                  <td>{fmt(p.extendedNetPrice)}</td>
                  <td>{fmt(p.extendedWeighted)}</td>
                  <td>{p.licenseModel}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );

  // ── Tab 3: Available Catalog ──────────────────────────────────────────────
  const renderAvailable = () => (
    <div className="tab-content">
      <section className="panel">
        <div className="section-heading">
          <h2>Available catalog (not yet deployed)</h2>
          <span>{analysis.availableProducts.length} items with no active deployment</span>
        </div>
        <p className="panel-note">Products where <strong>Deployment Plan Quantity = 0</strong> — entitlements owned but not yet deployed. These represent unrealised value opportunities.</p>
        <div className="table-scroll">
          <table className="priority-table">
            <thead>
              <tr>
                <th>Licence Part</th>
                <th>Description</th>
                <th>Platform</th>
                <th>IBM Cost (unit)</th>
                <th>Client Price (unit)</th>
                <th>License Model</th>
                <th>Category</th>
              </tr>
            </thead>
            <tbody>
              {analysis.availableProducts.map((p, i) => (
                <tr key={`avail-${i}`}>
                  <td><code>{p.licensePart}</code></td>
                  <td>{p.description}</td>
                  <td><span className={`badge badge-platform-${(p.platform||'').toLowerCase()}`}>{p.platform}</span></td>
                  <td>{fmt(p.ibmCost)}</td>
                  <td>{fmt(p.clientPrice)}</td>
                  <td>{p.licenseModel}</td>
                  <td>{p.category}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );

  // ── Tab 4: Client Presentations ───────────────────────────────────────────
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
                    <th>Phase</th>
                    <th>Key Activities</th>
                    <th>Owner(s)</th>
                    <th>Start</th>
                    <th>Duration</th>
                    <th>Milestones / Deliverables</th>
                  </tr>
                </thead>
                <tbody>
                  {llmRows.map((row, i) => (
                    <tr key={`llm-row-${i}`}>
                      <td>{row.week}</td>
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

  // ── Tab 7: Licensing & Cost ───────────────────────────────────────────────
  const renderLicensing = () => (
    <div className="tab-content">
      {/* Full catalog pricing table */}
      <section className="panel">
        <div className="section-heading">
          <h2>Full catalog — pricing overview</h2>
          <span>{analysis.products.length} line items (top 20 by client price shown)</span>
        </div>
        <div className="table-scroll">
          <table className="priority-table">
            <thead>
              <tr>
                <th>Licence Part</th>
                <th>Product</th>
                <th>Platform</th>
                <th>Status</th>
                <th>Category</th>
                <th>License Model</th>
                <th>IBM Cost (unit)</th>
                <th>Client Price (unit)</th>
              </tr>
            </thead>
            <tbody>
              {analysis.deploymentPriorities.map((item, i) => (
                <tr key={`cat-${i}`} className={item.isDeployed ? 'row-deployed' : ''}>
                  <td><code>{item.licensePart}</code></td>
                  <td>{item.product}</td>
                  <td><span className={`badge badge-platform-${(item.platform||'').toLowerCase()}`}>{item.platform}</span></td>
                  <td>
                    <span className={item.isDeployed ? 'badge badge-deployed' : 'badge badge-available'}>
                      {item.isDeployed ? `Deployed (${fmtQty(item.deploymentQty)})` : 'Available'}
                    </span>
                  </td>
                  <td>{item.category}</td>
                  <td>{item.licenseModel}</td>
                  <td>{fmt(item.ibmCost)}</td>
                  <td>{fmt(item.clientPrice)}</td>
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
                {item.isDeployed && <span className="badge badge-deployed">Deployed</span>}
              </div>
              <p className="recommendation-title">{item.title}</p>
              <p>{item.insight}</p>
              <div className="price-row">
                <span>IBM cost: <strong>{fmt(item.ibmCost)}</strong></span>
                <span>Client price: <strong>{fmt(item.clientPrice)}</strong></span>
              </div>
              <small>Part: {item.ssPart}</small>
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
    overview:      renderOverview,
    deployed:      renderDeployed,
    available:     renderAvailable,
    presentations: renderPresentations,
    deployment:    renderDeploymentPlan,
    plans:         renderPlans,
    licensing:     renderLicensing,
    signals:       renderSignals,
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
