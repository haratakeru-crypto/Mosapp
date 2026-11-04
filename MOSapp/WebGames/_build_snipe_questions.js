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
      if (c === '"' && src[i + 1] === '"') { cur += '"'; i++; }
      else if (c === '"') q = false;
      else cur += c;
    } else if (c === '"') q = true;
    else if (c === ",") { row.push(cur); cur = ""; }
    else if (c === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; }
    else if (c !== "\r") cur += c;
  }
  if (cur.length || row.length) { row.push(cur); rows.push(row); }
  return rows.filter((r) => r.some((c) => String(c).trim() !== ""));
}

function hash(id) {
  return String(id).split("-").reduce((n, part) => n + (Number(part) || 0), 0);
}

function build(csvPath, comment) {
  const rows = parseCsv(fs.readFileSync(csvPath, "utf8")).slice(1);
  const buttons = [];
  rows.forEach((r) => {
    const button = (r[7] || "").trim();
    if (button && !buttons.includes(button)) buttons.push(button);
  });
  const cards = [];
  rows.forEach((r) => {
    const id = r[0];
    const cat = r[1];
    const q = r[2];
    const correct = r[3];
    const wrong = r[4];
    const button = (r[7] || "").trim();
    cards.push({ id: id + "-tab", sourceId: id, cat, q, correct, wrong });
    if (!button || button === correct) return;
    const others = buttons.filter((b) => b !== button);
    if (!others.length) return;
    cards.push({
      id: id + "-btn",
      sourceId: id,
      cat,
      q,
      correct: button,
      wrong: others[hash(id) % others.length]
    });
  });
  const out = comment + "\nconst QUESTIONS = " + JSON.stringify(cards, null, 2) + ";\n";
  return { cards, out };
}

const jobs = [
  ["excel-3sec-snipe", "// Excel 演習。タブカード（-tab）と、H列がタブと違うときのボタンカード（-btn）。"],
  ["word-3sec-snipe", "// Word 演習。タブカード（-tab）と、H列がタブと違うときのボタンカード（-btn）。"],
  ["ppt-3sec-snipe", "// PowerPoint 演習。タブカード（-tab）と、H列がタブと違うときのボタンカード（-btn）。"]
];

jobs.forEach(([dir, comment]) => {
  const { cards, out } = build(path.join(__dirname, dir, "choices-review.csv"), comment);
  fs.writeFileSync(path.join(__dirname, dir, "questions.js"), out, "utf8");
  const tabs = cards.filter((c) => c.id.endsWith("-tab")).length;
  const btns = cards.filter((c) => c.id.endsWith("-btn")).length;
  const same = cards.filter((c) => c.correct === c.wrong);
  console.log(dir, "cards", cards.length, "tab", tabs, "btn", btns, "same", same.length);
  console.log("  sample", cards[0].id, cards[0].correct, "vs", cards[0].wrong);
  const btn = cards.find((c) => c.id.endsWith("-btn"));
  if (btn) console.log("  btn", btn.id, btn.correct, "vs", btn.wrong, "src", btn.sourceId);
});
