import { readFileSync } from "fs";
const sheet = JSON.parse(readFileSync(".bob/tmp/refcat/Weighting.json", "utf-8"));
const allRows = sheet.rows;

// Find header row index
let headerRowIdx = -1;
for (let i = 0; i < allRows.length; i++) {
  const row = allRows[i];
  const joined = row.map(v => (v ?? "").toString()).join("|");
  if (joined.includes("Deployment Plan Quantity") && joined.includes("Licence Description")) {
    headerRowIdx = i;
    break;
  }
}

// Find the longest header row - maybe there's a second part row with more columns
// Check rows around headerRowIdx
for (let i = headerRowIdx - 2; i <= headerRowIdx + 2; i++) {
  if (i >= 0 && i < allRows.length) {
    console.log(`Row ${i} (len=${allRows[i].length}):`, JSON.stringify(allRows[i]));
  }
}

// Also show all unique column values from a full data row
console.log("\nFirst full data row:");
const firstData = allRows[headerRowIdx + 1];
console.log(`Length: ${firstData.length}`);
firstData.forEach((v, i) => console.log(`  [${i}]: ${JSON.stringify(v)}`));
