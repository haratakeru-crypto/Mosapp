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
  return rows;
}

function writeSnipeJs(csvPath, outPath, comment) {
  const rows = parseCsv(fs.readFileSync(csvPath, "utf8")).slice(1).filter((r) => r[0]);
  const qs = rows.map((r) => ({
    id: r[0],
    cat: r[1],
    q: r[2],
    correct: r[3],
    wrong: r[4]
  }));
  const body = JSON.stringify(qs, null, 2);
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, `${comment}\nconst QUESTIONS = ${body};\n`, "utf8");
  console.log("snipe", outPath, qs.length);
}

function writeJudgeJs(jsonPath, outPath, cats, comment) {
  const data = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
  const lines = [comment, "const QUESTIONS = ["];
  let prev = null;
  data.tasks.forEach((t, i) => {
    if (prev !== null && prev !== t.projectId) lines.push("");
    prev = t.projectId;
    const q = t.q.replace(/\r\n/g, "\n").replace(/\n+/g, " ").replace(/\s+/g, " ").trim();
    const cat = cats[t.projectId] || "";
    const comma = i === data.tasks.length - 1 ? "" : ",";
    lines.push(`  { id: "${t.projectId}-${t.taskId}", cat: ${JSON.stringify(cat)}, q: ${JSON.stringify(q)} }${comma}`);
  });
  lines.push("];");
  lines.push("");
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, lines.join("\n"), "utf8");
  console.log("judge", outPath, data.tasks.length);
}

function copyHtml(src, dest, fromApp, toApp, kindTitle) {
  let text = fs.readFileSync(src, "utf8");
  text = text.split(fromApp).join(toApp);
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.writeFileSync(dest, text, "utf8");
  console.log("html", dest, kindTitle);
}

const WORD_JUDGE_CATS = {
  1: "保存とページ", 2: "書式と挿入", 3: "段組みと校閲", 4: "デザインとコメント",
  5: "表と図", 6: "リンクと図形", 7: "校閲と図形", 8: "挿入と参考資料"
};
const PPT_JUDGE_CATS = {
  1: "スライド編集", 2: "スライドと切り替え", 3: "レイアウトと図",
  4: "アニメーションと表", 5: "ズームと配布", 6: "マスターとメディア"
};

writeSnipeJs(
  path.join(__dirname, "word-3sec-snipe/choices-review.csv"),
  path.join(__dirname, "word-3sec-snipe/questions.js"),
  "// Word 演習タブの 61 問。correct は解答手順の最初のタブ/ボタン。"
);
writeSnipeJs(
  path.join(__dirname, "ppt-3sec-snipe/choices-review.csv"),
  path.join(__dirname, "ppt-3sec-snipe/questions.js"),
  "// PowerPoint 演習タブの 68 問。correct は解答手順の最初のタブ/ボタン。"
);
writeJudgeJs(
  path.join(__dirname, "_word_ruitai.json"),
  path.join(__dirname, "word-3sec-judge/questions.js"),
  WORD_JUDGE_CATS,
  "// Word アプリの類題タブ（類題セット1、プロジェクト1〜8）の各タスクを元にしたジャッジ問題。"
);
writeJudgeJs(
  path.join(__dirname, "_ppt_ruitai.json"),
  path.join(__dirname, "ppt-3sec-judge/questions.js"),
  PPT_JUDGE_CATS,
  "// PowerPoint アプリの類題タブ（類題セット1、プロジェクト1〜6）の各タスクを元にしたジャッジ問題。"
);

copyHtml(
  path.join(__dirname, "excel-3sec-judge/index.html"),
  path.join(__dirname, "word-3sec-judge/index.html"),
  "Excel", "Word", "judge-word"
);
copyHtml(
  path.join(__dirname, "excel-3sec-judge/index.html"),
  path.join(__dirname, "ppt-3sec-judge/index.html"),
  "Excel", "PowerPoint", "judge-ppt"
);
copyHtml(
  path.join(__dirname, "excel-3sec-snipe/index.html"),
  path.join(__dirname, "word-3sec-snipe/index.html"),
  "Excel", "Word", "snipe-word"
);
copyHtml(
  path.join(__dirname, "excel-3sec-snipe/index.html"),
  path.join(__dirname, "ppt-3sec-snipe/index.html"),
  "Excel", "PowerPoint", "snipe-ppt"
);
