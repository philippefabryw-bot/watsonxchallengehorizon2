import { readFileSync } from "fs";

const sheet = JSON.parse(readFileSync(".bob/tmp/refcat/Weighting.json", "utf-8"));
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

if (headerRowIdx === -1) { console.error("Cannot find header row"); process.exit(1); }

const headers = allRows[headerRowIdx].map(v => (v ?? "").toString().trim());
const dataRows = allRows.slice(headerRowIdx + 1);
const idx = {};
headers.forEach((h, i) => { if (h) idx[h] = i; });

const getNum = (row, col) => { const i = idx[col]; if (i===undefined) return null; const v = row[i]; if (v===null||v===undefined||v==="#UNCALCULATED") return null; const n=Number(v); return isNaN(n)?null:n; };
const getStr = (row, col) => { const i = idx[col]; if (i===undefined) return null; const v = row[i]; return (v!==null&&v!==undefined&&v!=="#UNCALCULATED") ? String(v).trim() : null; };

const results = [];
for (const row of dataRows) {
  const description = getStr(row, "Licence Description");
  if (!description) continue;
  const licPart = getStr(row, "Licence Part No");
  const ssPart  = getStr(row, "S&S Part No");
  if (!licPart && !ssPart) continue;
  const deployQty = getNum(row, "Deployment Plan Quantity");
  results.push({
    licencePart:      licPart,
    ssPart:           ssPart,
    description:      description,
    deploymentQty:    deployQty,
    isDeployed:       deployQty !== null && deployQty > 0,
    platform:         getStr(row, "z"),
    licenceEntitled:  getNum(row, "License Entitled Price"),
    ssEntitled:       getNum(row, "S&S Entitled Price"),
    drawdownEntitled: getNum(row, "Drawdown Price - Entitled"),
    discount:         getNum(row, "Discount"),
    ibmCost:          getNum(row, "Drawdown Price - Discounted"),
    clientPrice:      getNum(row, "Drawdown Price - Weighted"),
    extendedNetPrice: getNum(row, "Extended Net Price"),
    extendedWeighted: getNum(row, "Extended Weighted Price"),
  });
}

console.log(JSON.stringify(results, null, 2));
