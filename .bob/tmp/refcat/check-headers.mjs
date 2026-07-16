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
const headers = allRows[headerRowIdx];
console.log(JSON.stringify({ headerRowIdx, headers }));
// Also show first 3 data rows
const d = allRows.slice(headerRowIdx+1, headerRowIdx+4);
console.log(JSON.stringify(d, null, 2));
