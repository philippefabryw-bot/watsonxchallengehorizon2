import { readFileSync } from "fs";

// All column positions are known from the initial read_xlsx sample
// Header row (Excel row 33, 0-based array index 14 in dump):
// [0]=null [1]=check [2]=S&S Part No [3]=Licence Part No [4]=Licence Description
// [5]=Deployment Plan Quantity [6]=z(platform) [7]=License Entitled Price
// [8]=S&S Entitled Price [9]=Drawdown Price - Entitled [10]=Discount
// [11]=License Unit Discounted Price [12]=S&S Unit Discounted Price
// [13]=Drawdown Price - Discounted (IBM cost) [14]=Extended Net Price [15]=null
// [16]=Drawdown Price - Weighted (client price) [17]=Extended Weighted Price
// [18]=null [19]=Check [20]=null
// [21]=Licence Part No(contract) [22]=S&S Part No(contract) [23]=Licence Description(contract)
// [24]=Unit Exchange Price [25]=Type [26]=null [27]=Qty at DMSS price [28]=Qty at Weighted price

const sheet = JSON.parse(readFileSync(".bob/tmp/refcat/Weighting.json", "utf-8"));
const allRows = sheet.rows;

// Find header row
let headerRowIdx = -1;
for (let i = 0; i < allRows.length; i++) {
  const r = allRows[i];
  if (r.length >= 5 && String(r[4] ?? "").includes("Licence Description") && String(r[5] ?? "").includes("Deployment")) {
    headerRowIdx = i;
    break;
  }
}
if (headerRowIdx === -1) { console.error("Header row not found"); process.exit(1); }

const dataRows = allRows.slice(headerRowIdx + 1);

// Positional column indices (confirmed from initial sample)
const C = {
  ssPart:            2,
  licPart:           3,
  description:       4,
  deployQty:         5,
  platform:          6,
  licEntitled:       7,
  ssEntitled:        8,
  drawdownEntitled:  9,
  discount:          10,
  licDiscounted:     11,
  ssDiscounted:      12,
  ibmCost:           13,   // Drawdown Price - Discounted
  extendedNet:       14,
  clientPrice:       16,   // Drawdown Price - Weighted
  extendedWeighted:  17,
};

const safeNum = (row, i) => {
  if (i >= row.length) return null;
  const v = row[i];
  if (v === null || v === undefined || v === "#UNCALCULATED") return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
};
const safeStr = (row, i) => {
  if (i >= row.length) return null;
  const v = row[i];
  if (v === null || v === undefined || v === "#UNCALCULATED") return null;
  return String(v).trim() || null;
};

const results = [];
for (const row of dataRows) {
  const description = safeStr(row, C.description);
  if (!description) continue;
  const licPart = safeStr(row, C.licPart);
  const ssPart  = safeStr(row, C.ssPart);
  if (!licPart && !ssPart) continue;

  const deployQty = safeNum(row, C.deployQty);
  const ibmCost   = safeNum(row, C.ibmCost);
  const clientPrice = safeNum(row, C.clientPrice);

  // Derive product family from description
  const desc = description.toLowerCase();
  const platform = safeStr(row, C.platform) || (desc.includes("for ibm z") || desc.includes("for z ") ? "z" : "distributed");

  results.push({
    licencePart:      licPart,
    ssPart:           ssPart,
    description:      description,
    deploymentQty:    deployQty,
    isDeployed:       deployQty !== null && deployQty > 0,
    platform:         platform,
    licenceEntitled:  safeNum(row, C.licEntitled),
    ssEntitled:       safeNum(row, C.ssEntitled),
    drawdownEntitled: safeNum(row, C.drawdownEntitled),
    discount:         safeNum(row, C.discount),
    ibmCost:          ibmCost,
    clientPrice:      clientPrice,
    extendedNetPrice: safeNum(row, C.extendedNet),
    extendedWeighted: safeNum(row, C.extendedWeighted),
  });
}

console.log(JSON.stringify(results, null, 2));
