import { useMemo, useState } from 'react';

const marketSignals = [
  {
    theme: 'AI governance and trusted AI',
    demand: 'High',
    summary:
      'Organizations are prioritizing governed GenAI rollouts, model risk controls, and explainability for regulated workloads.',
    drivers: ['EU AI Act readiness', 'Board-level risk oversight', 'Need for auditable copilots'],
    recommendedCatalogTags: ['ai', 'governance', 'security', 'automation'],
  },
  {
    theme: 'Application modernization and hybrid integration',
    demand: 'High',
    summary:
      'Enterprises continue moving legacy estates to container, API, and event-driven architectures while protecting core systems.',
    drivers: ['Cloud cost optimization', 'Faster release cycles', 'Mainframe and distributed coexistence'],
    recommendedCatalogTags: ['integration', 'api', 'container', 'automation'],
  },
  {
    theme: 'Data fabric and observability',
    demand: 'Medium',
    summary:
      'Data leaders are investing in governed data access, lineage, and monitoring to improve AI outcomes and reduce operational blind spots.',
    drivers: ['Need for trusted data products', 'Platform engineering adoption', 'Compliance reporting pressure'],
    recommendedCatalogTags: ['data', 'observability', 'governance', 'analytics'],
  },
  {
    theme: 'Zero trust security modernization',
    demand: 'High',
    summary:
      'Security investments are shifting toward identity-aware controls, posture management, and automated incident response.',
    drivers: ['Ransomware pressure', 'Distributed workforce', 'Software supply chain exposure'],
    recommendedCatalogTags: ['security', 'identity', 'automation', 'compliance'],
  },
];

const licensePatterns = [
  {
    name: 'Perpetual or fixed-capacity licensing',
    keywords: ['perpetual', 'pvus', 'processor', 'socket', 'fixed capacity'],
    recommendation:
      'Consider subscription or consumption-aligned packaging so clients can scale licenses with usage and hybrid deployments.',
  },
  {
    name: 'User-based licensing',
    keywords: ['authorized user', 'user', 'seat'],
    recommendation:
      'Evaluate role-based tiers, API consumption, or enterprise bundles to align better with automation and machine-to-machine workloads.',
  },
  {
    name: 'Support-only renewals',
    keywords: ['support renewal', 'subscription & support', 's&s'],
    recommendation:
      'Package modernization services, managed operations, or AI accelerators with renewals to create stronger upgrade incentives.',
  },
];

const sampleCatalog = `Product,Category,LicenseModel,Deployment,Capabilities,Version,LastUpdated
IBM API Connect,Integration,PVU subscription,Hybrid cloud,API management;developer portal;security,10.0.8,2023-11-15
IBM MQ,Integration,Perpetual + S&S,On-prem and cloud,messaging;event streaming integration,9.3.4,2022-09-10
IBM Security Verify,Security,User-based SaaS,SaaS,identity;access management;MFA,11.0,2024-03-04
IBM watsonx.ai,AI,Consumption,SaaS and hybrid,foundation models;prompt lab;governance,1.1.2,2024-05-22
IBM DataStage,Data,Authorized user + PVU,Hybrid,data integration;ETL;data quality,11.7,2021-06-30`;

const parseCatalog = (text) => {
  const rows = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (rows.length < 2) {
    return [];
  }

  const headers = rows[0].split(',').map((header) => header.trim());

  return rows.slice(1).map((row, index) => {
    const values = row.split(',').map((value) => value.trim());
    const item = headers.reduce((accumulator, header, columnIndex) => {
      accumulator[header] = values[columnIndex] ?? '';
      return accumulator;
    }, {});

    return {
      id: `${item.Product || 'product'}-${index}`,
      ...item,
    };
  });
};

const getProductTags = (product) =>
  [product.Category, product.Deployment, product.Capabilities, product.LicenseModel]
    .join(';')
    .toLowerCase()
    .split(/[^a-z0-9+]+/)
    .filter(Boolean);

const getMonthsOld = (dateText) => {
  const parsed = new Date(dateText);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  const now = new Date();
  return (now.getFullYear() - parsed.getFullYear()) * 12 + (now.getMonth() - parsed.getMonth());
};

const analyzeCatalog = (products) => {
  const licenseRecommendations = [];
  const refreshRecommendations = [];
  const useCaseRecommendations = [];

  products.forEach((product) => {
    const searchable = `${product.LicenseModel} ${product.Capabilities} ${product.Deployment}`.toLowerCase();
    const tags = getProductTags(product);
    const ageInMonths = getMonthsOld(product.LastUpdated);

    licensePatterns.forEach((pattern) => {
      if (pattern.keywords.some((keyword) => searchable.includes(keyword))) {
        licenseRecommendations.push({
          product: product.Product,
          title: pattern.name,
          insight: pattern.recommendation,
          rationale: `Detected ${pattern.keywords.filter((keyword) => searchable.includes(keyword)).join(', ')} in the current packaging.`,
        });
      }
    });

    if (ageInMonths !== null && ageInMonths >= 12) {
      refreshRecommendations.push({
        product: product.Product,
        title: 'Catalog refresh recommended',
        insight:
          'Update messaging, release details, and deployment guidance to reflect current capabilities and competitive positioning.',
        rationale: `Catalog metadata appears approximately ${ageInMonths} months old.`,
      });
    }

    marketSignals.forEach((signal) => {
      const overlap = signal.recommendedCatalogTags.filter((tag) => tags.includes(tag));
      if (overlap.length > 0) {
        useCaseRecommendations.push({
          product: product.Product,
          title: signal.theme,
          insight: `Position ${product.Product} for ${signal.theme.toLowerCase()} by emphasizing ${overlap.join(', ')} strengths already present in the catalog.`,
          rationale: signal.summary,
          drivers: signal.drivers,
        });
      }
    });
  });

  return {
    licenseRecommendations,
    refreshRecommendations,
    useCaseRecommendations,
    portfolioHighlights: {
      totalProducts: products.length,
      categories: [...new Set(products.map((product) => product.Category).filter(Boolean))],
      deploymentMix: [...new Set(products.map((product) => product.Deployment).filter(Boolean))],
    },
  };
};

function App() {
  const [catalogText, setCatalogText] = useState(sampleCatalog);

  const products = useMemo(() => parseCatalog(catalogText), [catalogText]);
  const analysis = useMemo(() => analyzeCatalog(products), [products]);

  const onFileUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    const text = await file.text();
    setCatalogText(text);
  };

  return (
    <div className="app-shell">
      <header className="hero">
        <div>
          <p className="eyebrow">IBM Software Portfolio Advisor</p>
          <h1>Catalog modernization recommendations in one workspace</h1>
          <p className="hero-copy">
            Upload a CSV export of an IBM software catalog to identify license modernization opportunities,
            stale catalog entries, and net-new use cases aligned to current market demand.
          </p>
        </div>
        <div className="hero-card">
          <h2>Expected columns</h2>
          <p>Product, Category, LicenseModel, Deployment, Capabilities, Version, LastUpdated</p>
          <label className="upload-button">
            Upload catalog CSV
            <input type="file" accept=".csv,text/csv" onChange={onFileUpload} />
          </label>
        </div>
      </header>

      <main className="layout">
        <section className="panel input-panel">
          <div className="section-heading">
            <h2>Catalog input</h2>
            <span>{products.length} products parsed</span>
          </div>
          <textarea
            value={catalogText}
            onChange={(event) => setCatalogText(event.target.value)}
            spellCheck="false"
            aria-label="Catalog CSV"
          />
        </section>

        <section className="panel summary-panel">
          <div className="section-heading">
            <h2>Portfolio snapshot</h2>
          </div>
          <div className="stats-grid">
            <article>
              <strong>{analysis.portfolioHighlights.totalProducts}</strong>
              <span>Products</span>
            </article>
            <article>
              <strong>{analysis.licenseRecommendations.length}</strong>
              <span>License actions</span>
            </article>
            <article>
              <strong>{analysis.refreshRecommendations.length}</strong>
              <span>Refresh flags</span>
            </article>
            <article>
              <strong>{analysis.useCaseRecommendations.length}</strong>
              <span>Use-case ideas</span>
            </article>
          </div>
          <div className="chips">
            {analysis.portfolioHighlights.categories.map((category) => (
              <span key={category} className="chip">
                {category}
              </span>
            ))}
          </div>
          <div className="chips deployment-chips">
            {analysis.portfolioHighlights.deploymentMix.map((deployment) => (
              <span key={deployment} className="chip chip-secondary">
                {deployment}
              </span>
            ))}
          </div>
        </section>

        <section className="panel recommendations-panel">
          <div className="section-heading">
            <h2>License modernization</h2>
          </div>
          {analysis.licenseRecommendations.length === 0 ? (
            <p className="empty-state">No license modernization patterns detected from the current catalog input.</p>
          ) : (
            analysis.licenseRecommendations.map((item) => (
              <article className="recommendation-card" key={`${item.product}-${item.title}`}>
                <h3>{item.product}</h3>
                <p className="recommendation-title">{item.title}</p>
                <p>{item.insight}</p>
                <small>{item.rationale}</small>
              </article>
            ))
          )}
        </section>

        <section className="panel recommendations-panel">
          <div className="section-heading">
            <h2>Catalog refresh</h2>
          </div>
          {analysis.refreshRecommendations.length === 0 ? (
            <p className="empty-state">No stale catalog entries were identified from the provided LastUpdated values.</p>
          ) : (
            analysis.refreshRecommendations.map((item) => (
              <article className="recommendation-card" key={`${item.product}-${item.title}`}>
                <h3>{item.product}</h3>
                <p className="recommendation-title">{item.title}</p>
                <p>{item.insight}</p>
                <small>{item.rationale}</small>
              </article>
            ))
          )}
        </section>

        <section className="panel recommendations-panel">
          <div className="section-heading">
            <h2>New use cases from market signals</h2>
          </div>
          {analysis.useCaseRecommendations.length === 0 ? (
            <p className="empty-state">No market-aligned use cases were matched. Add richer capability keywords to expand recommendations.</p>
          ) : (
            analysis.useCaseRecommendations.map((item, index) => (
              <article className="recommendation-card" key={`${item.product}-${index}`}>
                <h3>{item.product}</h3>
                <p className="recommendation-title">{item.title}</p>
                <p>{item.insight}</p>
                <small>{item.rationale}</small>
                <ul>
                  {item.drivers.map((driver) => (
                    <li key={driver}>{driver}</li>
                  ))}
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
      </main>
    </div>
  );
}

export default App;
