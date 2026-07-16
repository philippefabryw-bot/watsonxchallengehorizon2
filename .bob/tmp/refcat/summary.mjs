import { readFileSync } from "fs";
const data = JSON.parse(readFileSync(".bob/tmp/refcat/catalog-data4.json", "utf-8"));
const deployed = data.filter(r => r.isDeployed);
const available = data.filter(r => !r.isDeployed);
console.error(`Total: ${data.length}, Deployed: ${deployed.length}, Available: ${available.length}`);

// Show some available items
const sample = available.slice(0, 5);
console.log("Available sample:", JSON.stringify(sample, null, 2));

// Show platform breakdown
const platforms = {};
data.forEach(r => { platforms[r.platform] = (platforms[r.platform]||0)+1; });
console.error("Platforms:", JSON.stringify(platforms));
