const fs = require("fs");
const path = require("path");
const root = path.join(__dirname, "..");

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cur = "";
  let q = false;
  const src = text.replace(/^\uFEFF/, "");
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (q) {
      if (c === '"' && src[i + 1] === '"') { cur += '"'; i++; }
      else if (c === '"') q = false;
      else cur += c;
    } else if (c === '"') q = true;
    else if (c === ",") { row.push(cur); cur = ""; }
    else if (c === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; }
    else if (c !== "\r") cur += c;
  }
  if (cur.length || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

function dump(csvPath, qCol, stepsCol, ids, pidCol, tidCol) {
  const rows = parseCsv(fs.readFileSync(csvPath, "utf8")).slice(1);
  const want = new Set(ids);
  for (const r of rows) {
    const id = `${r[pidCol]}-${r[tidCol]}`;
    if (want.has(id)) {
      console.log("====", id);
      console.log("Q:", (r[qCol] || "").replace(/\s+/g, " ").slice(0, 120));
      console.log("STEPS:", (r[stepsCol] || "").replace(/\r\n/g, " / ").slice(0, 400));
      console.log("");
    }
  }
}

dump(
  path.join(root, "MOS Word app/Reference/CSV/MOS模擬アプリ正誤判定表251120_.csv"),
  2, 3, ["4-5", "4-6", "8-1", "10-3", "2-4", "6-4", "9-1"], 0, 1
);
console.log("----- PPT -----");
dump(
  path.join(root, "../task/PP修正版問題文一覧.csv"),
  3, 4, ["1-7", "3-1", "3-5", "4-1", "5-7", "9-1", "9-2", "9-3", "9-4", "10-1"], 1, 2
);
