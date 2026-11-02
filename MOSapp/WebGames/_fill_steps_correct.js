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
      if (c === '"' && src[i + 1] === '"') {
        cur += '"';
        i++;
      } else if (c === '"') q = false;
      else cur += c;
    } else if (c === '"') q = true;
    else if (c === ",") {
      row.push(cur);
      cur = "";
    } else if (c === "\n") {
      row.push(cur);
      rows.push(row);
      row = [];
      cur = "";
    } else if (c !== "\r") cur += c;
  }
  if (cur.length || row.length) {
    row.push(cur);
    rows.push(row);
  }
  return rows.filter((r) => r.some((c) => String(c).trim() !== ""));
}

function csvEscape(s) {
  return `"${String(s).replace(/"/g, '""')}"`;
}

function splitSteps(raw) {
  const s = String(raw || "").replace(/\r\n/g, "\n");
  const parts = [];
  let cur = "";
  let quote = null;
  const pairs = { "「": "」", "『": "』", "\"": "\"", "“": "”" };
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (quote) {
      cur += c;
      if (c === quote) quote = null;
      continue;
    }
    if (pairs[c]) {
      quote = pairs[c];
      cur += c;
      continue;
    }
    if (c === "↓") {
      parts.push(cur);
      cur = "";
      continue;
    }
    cur += c;
  }
  if (cur.trim()) parts.push(cur);
  return parts;
}

function cleanPart(s) {
  let part = String(s || "")
    .replace(/\[/g, "［")
    .replace(/\]/g, "］")
    .replace(/\s+/g, "")
    .trim();
  if (/］タブ/.test(part) && !part.includes("［")) part = "［" + part;
  return part;
}

function isLeadingSkip(part) {
  const p = part.replace(/\s+/g, "");
  if (/タブ/.test(p)) return false;
  if (/［/.test(p)) return false;
  if (/入力|右クリック/.test(p)) return false;
  return /選択|クリック|カーソル|ドラッグ/.test(p);
}

function stepsCorrect(raw) {
  const parts = splitSteps(raw).map(cleanPart).filter(Boolean);
  let i = 0;
  while (i < parts.length && isLeadingSkip(parts[i])) i++;
  const kept = (i < parts.length ? parts.slice(i) : parts).filter(Boolean);
  return kept.join("→");
}

function indexById(rows, idA, idB, stepsCol) {
  const map = new Map();
  rows.slice(1).forEach((r) => {
    const id = `${String(r[idA]).trim()}-${String(r[idB]).trim()}`;
    map.set(id, r[stepsCol] || "");
  });
  return map;
}

function fill(reviewPath, stepsMap) {
  const rows = parseCsv(fs.readFileSync(reviewPath, "utf8"));
  const header = rows[0].slice(0, 6);
  header[5] = header[5] || "steps_note";
  if (header.length < 6) throw new Error("expected 6 columns in " + reviewPath);
  const out = [["id", "cat", "q", "correct", "wrong", "steps_note", "steps_correct"]];
  const missing = [];
  rows.slice(1).forEach((r) => {
    const id = r[0];
    const steps = stepsMap.get(id);
    if (steps == null) missing.push(id);
    const g = stepsCorrect(steps || "");
    if (!g) missing.push(id + " empty");
    out.push([r[0], r[1], r[2], r[3], r[4], r[5], g]);
  });
  const text = "\uFEFF" + out.map((r) => r.map(csvEscape).join(",")).join("\r\n") + "\r\n";
  fs.writeFileSync(reviewPath, text, "utf8");
  return { count: out.length - 1, missing, sample: out.slice(1, 8).map((r) => r[0] + "\t" + r[6]) };
}

const wordSteps = indexById(
  parseCsv(fs.readFileSync(path.join(root, "MOS Word app/Reference/CSV/MOS模擬アプリ正誤判定表251120_.csv"), "utf8")),
  0, 1, 3
);
const pptSteps = indexById(
  parseCsv(fs.readFileSync(path.join(root, "../task/PP修正版問題文一覧.csv"), "utf8")),
  1, 2, 4
);

const word = fill(path.join(__dirname, "word-3sec-snipe/choices-review.csv"), wordSteps);
const ppt = fill(path.join(__dirname, "ppt-3sec-snipe/choices-review.csv"), pptSteps);
console.log("WORD", word.count, "missing", word.missing.join(","));
console.log(word.sample.join("\n"));
console.log("PPT", ppt.count, "missing", ppt.missing.join(","));
console.log(ppt.sample.join("\n"));

function show(file, ids) {
  const rows = parseCsv(fs.readFileSync(file, "utf8"));
  ids.forEach((id) => {
    const r = rows.find((x) => x[0] === id);
    console.log("----", id);
    console.log(r ? r[6] : "NOT FOUND");
  });
}
function audit(label, file) {
  const rows = parseCsv(fs.readFileSync(file, "utf8"));
  console.log(label, "cols", rows[0].join("|"), "rows", rows.length - 1);
  rows.slice(1).forEach((r) => {
    const g = r[6] || "";
    if (!g.startsWith("［") && !/入力|右クリック/.test(g.slice(0, 24))) {
      console.log("START", r[0], g.slice(0, 80));
    }
  });
}
audit("WORD", path.join(__dirname, "word-3sec-snipe/choices-review.csv"));
audit("PPT", path.join(__dirname, "ppt-3sec-snipe/choices-review.csv"));
show(path.join(__dirname, "word-3sec-snipe/choices-review.csv"), ["1-2", "2-1", "2-4", "4-2", "4-5", "8-1", "9-1", "10-3"]);
show(path.join(__dirname, "ppt-3sec-snipe/choices-review.csv"), ["1-1", "1-7", "3-5", "4-1", "5-7", "8-1", "9-1", "9-4", "9-5"]);
