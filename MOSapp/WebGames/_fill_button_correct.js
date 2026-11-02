const fs = require("fs");
const path = require("path");

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
  return `"${String(s == null ? "" : s).replace(/"/g, '""')}"`;
}

function excelStepsCorrect(note) {
  const parts = String(note || "").split("→").map((s) => s.trim()).filter(Boolean);
  let i = 0;
  while (i < parts.length && !/［/.test(parts[i]) && /選択|あと/.test(parts[i])) i++;
  return parts.slice(i).join("→");
}

function buttonCorrect(steps) {
  const re = /［([^］]+)］/g;
  let m;
  const text = String(steps || "");
  while ((m = re.exec(text))) {
    const after = text.slice(m.index + m[0].length, m.index + m[0].length + 2);
    if (after.startsWith("タブ") || after.startsWith("キー")) continue;
    return "［" + m[1] + "］";
  }
  return "";
}

function writeRows(file, rows) {
  const text = "\uFEFF" + rows.map((r) => r.map(csvEscape).join(",")).join("\r\n") + "\r\n";
  const target = file;
  try {
    fs.writeFileSync(target, text, "utf8");
  } catch (e) {
    if (e.code !== "EBUSY") throw e;
    fs.writeFileSync(target + ".tmp", text, "utf8");
    console.log("LOCKED", target);
  }
}

function fillExcel(file) {
  const rows = parseCsv(fs.readFileSync(file, "utf8"));
  const out = [["id", "cat", "q", "correct", "wrong", "steps_note", "steps_correct", "button_correct"]];
  rows.slice(1).forEach((r) => {
    const g = excelStepsCorrect(r[5]);
    out.push([r[0], r[1], r[2], r[3], r[4], r[5], g, buttonCorrect(g)]);
  });
  writeRows(file, out);
  return out;
}

function fillFromG(file) {
  const rows = parseCsv(fs.readFileSync(file, "utf8"));
  const out = [["id", "cat", "q", "correct", "wrong", "steps_note", "steps_correct", "button_correct"]];
  rows.slice(1).forEach((r) => {
    const g = r[6] || "";
    out.push([r[0], r[1], r[2], r[3], r[4], r[5], g, buttonCorrect(g)]);
  });
  writeRows(file, out);
  return out;
}

const excel = fillExcel(path.join(__dirname, "excel-3sec-snipe/choices-review.csv"));
const word = fillFromG(path.join(__dirname, "word-3sec-snipe/choices-review.csv"));
const ppt = fillFromG(path.join(__dirname, "ppt-3sec-snipe/choices-review.csv"));

function show(label, rows, ids) {
  console.log("====", label, rows.length - 1);
  ids.forEach((id) => {
    const r = rows.find((x) => x[0] === id);
    console.log(id);
    console.log("  G", r[6]);
    console.log("  H", r[7]);
  });
  const empty = rows.slice(1).filter((r) => !r[7]).map((r) => r[0]);
  console.log("empty H", empty.join(","));
}

show("EXCEL", excel, ["1-1", "1-4", "2-7", "5-3", "8-1", "9-1"]);
show("WORD", word, ["1-1", "1-2", "2-4", "2-5", "3-1"]);
show("PPT", ppt, ["1-1", "1-2", "4-1", "9-1"]);
