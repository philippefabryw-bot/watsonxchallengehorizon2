import { readFileSync } from "fs";

// Read the raw Weighting sheet JSON (auto-dumped without correct header row)
// Real headers are on Excel row 20 (0-based index 19 in the sheet rows array from read_xlsx with header_row=20)
// We re-read the raw dump and manually skip to the real header row.

const DUMP_DIR = ".bob/tmp/xlsx-dumps/reference catalog-11e753e00cbb4c7a";
const sheet = JSON.parse(readFileSync(`${DUMP_DIR}/Weighting.json`, "utf-8"));

// The dump was taken without specifying header_row so all rows are in sheet.rows
// We need to find the real header row: it contains "Licence Description" or "Deployment Plan Quantity"
const allRows = sheet.rows;

let headerRowIdx = -1;
for (let i = 0; i < allRows.length; i++) {
  const row = allRows[i];
  const joined = row.map(v => (v ?? "").toString()).join("|");
  if (joined.includes("Deployment Plan Quantity") && joined.includes("Licence Description")) {
    headerRowIdx = i;
    break;
  }
}

if (headerRowIdx === -1) {
  console.error("Could not find header row");
  process.exit(1);
}

const headers = allRows[headerRowIdx].map(v => (v ?? "").toString().trim());
const dataRows = allRows.slice(headerRowIdx + 1);

// Build index map
const idx = {};
headers.forEach((h, i) => { if (h) idx[h] = i; });

// Column keys we care about
const COL_SS_PART      = "S&S Part No";
const COL_LIC_PART     = "Licence Part No";
const COL_DESCRIPTION  = "Licence Description";
const COL_DEPLOY_QTY   = "Deployment Plan Quantity";
const COL_PLATFORM     = "z";                         // PA or z
const COL_LIC_ENTITLED = "License Entitled Price";
const COL_SS_ENTITLED  = "S&S Entitled Price";
const COL_DRAWDOWN_ENTITLED = "Drawdown Price - Entitled";
const COL_DISCOUNT     = "Discount";
const COL_LIC_DISC     = "License Unit Discounted Price";
const COL_SS_DISC      = "S&S Unit Discounted Price";
const COL_DRAWDOWN_DISC   = "Drawdown Price - Discounted";  // IBM cost
const COL_DRAWDOWN_WTGT   = "Drawdown Price - Weighted";    // client price
const COL_EXT_NET      = "Extended Net Price";
const COL_EXT_WEIGHTED = "Extended Weighted Price";

const get = (row, colName) => {
  const i = idx[colName];
  if (i === undefined) return null;
  const v = row[i];
  if (v === null || v === undefined || v === "#UNCALCULATED") return null;
  return v;
};

const getNum = (row, colName) => {
  const v = get(row, colName);
  const n = Number(v);
  return isNaN(n) ? null : n;
};

const getString = (row, colName) => {
  const v = get(row, colName);
  return v !== null ? String(v).trim() : null;
};

const results = [];

for (const row of dataRows) {
  const description = getString(row, COL_DESCRIPTION);
  // Skip rows with no description or all-null
  if (!description) continue;

  const licPart = getString(row, COL_LIC_PART);
  const ssPart  = getString(row, COL_SS_PART);

  // Skip rows that look like template blank rows (no part number either)
  if (!licPart && !ssPart) continue;

  const deployQty       = getNum(row, COL_DEPLOY_QTY);
  const platform        = getString(row, COL_PLATFORM);
  const licEntitled     = getNum(row, COL_LIC_ENTITLED);
  const ssEntitled      = getNum(row, COL_SS_ENTITLED);
  const drawdownEntitled= getNum(row, COL_DRAWDOWN_ENTITLED);
  const discount        = getNum(row, COL_DISCOUNT);
  const drawdownDisc    = getNum(row, COL_DRAWDOWN_DISC);   // IBM cost
  const drawdownWeighted= getNum(row, COL_DRAWDOWN_WTGT);   // client price
  const extNet          = getNum(row, COL_EXT_NET);
  const extWeighted     = getNum(row, COL_EXT_WEIGHTED);

  results.push({
    licencePart:        licPart,
    ssPart:             ssPart,
    description:        description,
    deploymentQty:      deployQty,        // non-null = deployed, null = available only
    isDeployed:         deployQty !== null && deployQty > 0,
    platform:           platform,         // "PA" or "z"
    licenceEntitled:    licEntitled,
    ssEntitled:         ssEntitled,
    drawdownEntitled:   drawdownEntitled,
    discount:           discount,
    ibmCost:            drawdownDisc,     // Drawdown Price - Discounted = IBM cost
    clientPrice:        drawdownWeighted, // Drawdown Price - Weighted   = client pays
    extendedNetPrice:   extNet,
    extendedWeighted:   extWeighted,
  });
}

console.log(JSON.stringify(results, null, 2));
