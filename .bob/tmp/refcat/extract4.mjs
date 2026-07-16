// Read the reference catalog using the xlsx package already in node_modules
import { readFileSync } from "fs";
import { createRequire } from "module";
const require = createRequire(import.meta.url);

// Use xlsx from the workspace root node_modules
process.chdir("C:/Users/DanielRosa-Yee/.bob/playground/watsonxchallengehorizon2");
const XLSX = require("xlsx");

const buf = readFileSync("reference catalog.xlsx");
const wb = XLSX.read(buf, { type: "buffer", cellDates: true });
const ws = wb.Sheets["Weighting"];

const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

// Find the row containing both "Licence Description" and "Deployment Plan Quantity"
let hdrIdx = -1;
for (let i = 0; i < raw.length; i++) {
  const joined = raw[i].map(v => String(v ?? "")).join("|");
  if (joined.includes("Licence Description") && joined.includes("Deployment Plan Quantity")) {
    hdrIdx = i;
    break;
  }
}
if (hdrIdx === -1) { console.error("Header row not found"); process.exit(1); }

const headers = raw[hdrIdx].map(v => String(v ?? "").trim());
const dataRows = raw.slice(hdrIdx + 1);

// Build column index map
const idx = {};
headers.forEach((h, i) => { if (h) idx[h] = i; });
console.error("Headers found:", Object.keys(idx).join(", "));

const safeNum = (row, col) => {
  const i = idx[col];
  if (i === undefined) return null;
  const v = row[i];
  if (v === null || v === undefined) return null;
  if (typeof v === "string" && (v.startsWith("#") || !v.trim())) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
};
const safeStr = (row, col) => {
  const i = idx[col];
  if (i === undefined) return null;
  const v = row[i];
  if (v === null || v === undefined) return null;
  if (typeof v === "string" && v.startsWith("#")) return null;
  const s = String(v).trim();
  return s || null;
};

const results = [];
for (const row of dataRows) {
  const description = safeStr(row, "Licence Description");
  if (!description) continue;
  const licPart = safeStr(row, "Licence Part No");
  const ssPart  = safeStr(row, "S&S Part No");
  if (!licPart && !ssPart) continue;

  const deployQty   = safeNum(row, "Deployment Plan Quantity");
  const ibmCost     = safeNum(row, "Drawdown Price - Discounted");
  const clientPrice = safeNum(row, "Drawdown Price - Weighted");

  results.push({
    licencePart:      licPart,
    ssPart:           ssPart,
    description:      description,
    deploymentQty:    deployQty,
    isDeployed:       deployQty !== null && deployQty > 0,
    platform:         safeStr(row, "z"),
    licenceEntitled:  safeNum(row, "License Entitled Price"),
    ssEntitled:       safeNum(row, "S&S Entitled Price"),
    drawdownEntitled: safeNum(row, "Drawdown Price - Entitled"),
    discount:         safeNum(row, "Discount"),
    ibmCost:          ibmCost,
    clientPrice:      clientPrice,
    extendedNetPrice: safeNum(row, "Extended Net Price"),
    extendedWeighted: safeNum(row, "Extended Weighted Price"),
  });
}

console.log(JSON.stringify(results, null, 2));
