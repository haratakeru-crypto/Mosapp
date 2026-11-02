const fs = require("fs");
const path = require("path");
function parseCsv(text) {
  const rows = []; let row = []; let cur = ""; let q = false;
  const src = text.replace(/^\uFEFF/, "");
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (q) { if (c === '"' && src[i + 1] === '"') { cur += '"'; i++; } else if (c === '"') q = false; else cur += c; }
    else if (c === '"') q = true;
    else if (c === ",") { row.push(cur); cur = ""; }
    else if (c === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; }
    else if (c !== "\r") cur += c;
  }
  if (cur.length || row.length) { row.push(cur); rows.push(row); }
  return rows;
}
const csv = parseCsv(fs.readFileSync(path.join(__dirname, "../task/PP修正版問題文一覧.csv"), "utf8"));
for (const r of csv) {
  if (r[1] === "8" && ["1","2","3"].includes(r[2])) {
    console.log("====", r[1]+"-"+r[2]);
    console.log((r[3]||"").replace(/\s+/g," ").slice(0,100));
    console.log((r[4]||"").replace(/\s+/g," / ").slice(0,250));
  }
}
const w = parseCsv(fs.readFileSync("word-3sec-snipe/choices-review.csv","utf8"));
const p = parseCsv(fs.readFileSync("ppt-3sec-snipe/choices-review.csv","utf8"));
console.log("word rows", w.length-1, "ppt rows", p.length-1);
console.log("word 10-3", w.find(r=>r[0]==="10-3"));
console.log("ppt 8-1", p.find(r=>r[0]==="8-1"));
console.log("ppt 3-5", p.find(r=>r[0]==="3-5"));
console.log("ppt 9-1", p.find(r=>r[0]==="9-1"));
