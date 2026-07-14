import { readFileSync } from "fs";

const DUMP_DIR = ".bob/tmp/xlsx-dumps/ibm-license-catalog.xlsx-d7fd53df3e76cbc4";
const { rows, headers } = JSON.parse(readFileSync(`${DUMP_DIR}/Table_1_ACP__Cad_.json`, "utf-8"));

const idx = Object.fromEntries(headers.map((h, i) => [h, i]));

// Derive product family from description prefix before first keyword
const productFamily = (desc) => {
  if (!desc) return "Unknown";
  const m = desc.match(/^IBM\s+([\w\s]+?)\s+(?:Cartridge|Authorized|Floating|Virtual|Non-Production|Processor|Resource|Managed|Software|Developer|Simultaneous|Per\s|per\s)/i);
  return m ? `IBM ${m[1].trim()}` : desc.split(" ").slice(0, 4).join(" ");
};

// Infer license model from description
const licenseModel = (desc) => {
  const d = (desc || "").toLowerCase();
  if (d.includes("authorized user")) return "Authorized User";
  if (d.includes("floating user")) return "Floating User";
  if (d.includes("virtual processor core") || d.includes("processor value unit") || d.includes("pvu")) return "PVU / VPC";
  if (d.includes("resource unit")) return "Resource Unit";
  if (d.includes("managed") || d.includes("per virtual server")) return "Per Server";
  if (d.includes("simultaneous session")) return "Session";
  if (d.includes("process") || d.includes("establishment")) return "Capacity";
  return "Other";
};

const sample = rows.slice(0, 10).map((row) => ({
  ssPart: row[idx["S&S part number"]],
  licensePart: row[idx["License part number"]],
  description: row[idx["Description"]],
  royalty: row[idx["Royalty"]],
  consumptionRate: row[idx["Consumption rate (CAD)"]],
  price2029: row[idx["ACP S&S Unit Reference Price for NTE for 2029 (CAD)"]],
  price2030: row[idx["ACP S&S Unit Reference Price for NTE for 2030 (CAD)"]],
  productFamily: productFamily(row[idx["Description"]]),
  licenseModel: licenseModel(row[idx["Description"]]),
}));

// Count license model distribution
const modelCounts = {};
for (const row of rows) {
  const m = licenseModel(row[idx["Description"]]);
  modelCounts[m] = (modelCounts[m] ?? 0) + 1;
}

// Count product family distribution (top 15)
const familyCounts = {};
for (const row of rows) {
  const f = productFamily(row[idx["Description"]]);
  familyCounts[f] = (familyCounts[f] ?? 0) + 1;
}
const topFamilies = Object.entries(familyCounts)
  .sort((a, b) => b[1] - a[1])
  .slice(0, 15)
  .map(([name, count]) => ({ name, count }));

console.log(JSON.stringify({ sample, modelCounts, topFamilies }, null, 2));
