const fs = require("fs");
const path = require("path");

const json = JSON.parse(fs.readFileSync(
  path.join(__dirname, "../../mos_xaml_app/References/JSON/MOS演習問題文一覧.json"),
  "utf8"
));
const csvPath = path.join(
  __dirname,
  "../../mos_xaml_app/References/Libs/Group1/Docs/MOS模擬アプリ正誤判定表250702.xlsx - Excel新教材.csv"
);
const csv = fs.readFileSync(csvPath, "utf8").replace(/^\uFEFF/, "");

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cur = "";
  let q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"' && text[i + 1] === '"') {
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
  return rows;
}

function norm(s) {
  return (s || "")
    .replace(/^あなたは[^\n]*\n/, "")
    .replace(/["“”'']/g, "")
    .replace(/[［\[]/g, "[")
    .replace(/[］\]]/g, "]")
    .replace(/[「」『』]/g, "")
    .replace(/してください。?/g, "")
    .replace(/。+$/g, "")
    .replace(/\s+/g, "")
    .trim();
}

const csvItems = parseCsv(csv).slice(1).filter((r) => r.length >= 7).map((r) => ({
  q: r[5],
  steps: r[6]
}));

const TABS = [
  { re: /テーブル\s*デザイン/, label: "［テーブルデザイン］タブ" },
  { re: /グラフのデザイン/, label: "［グラフのデザイン］タブ" },
  { re: /ページ\s*レイアウト/, label: "［ページレイアウト］タブ" },
  { re: /ホーム/, label: "［ホーム］タブ" },
  { re: /挿入/, label: "［挿入］タブ" },
  { re: /数式/, label: "［数式］タブ" },
  { re: /データ/, label: "［データ］タブ" },
  { re: /校閲/, label: "［校閲］タブ" },
  { re: /表示/, label: "［表示］タブ" },
  { re: /ファイル/, label: "［ファイル］タブ" },
  { re: /書式/, label: "［書式］タブ" }
];

function extractFirst(steps) {
  const parts = (steps || "").split("↓").map((s) => s.trim()).filter(Boolean);
  const skip = /シートを選択|シートを選択|シート］を選択|シート」を選択|シート\[|シート［|セル範囲|セル【|セル「|任意のセル|グラフエリア|行番号|右クリック|問題のURL|数式バー/;
  for (const p of parts) {
    if (/名前ボックス/.test(p)) return { label: "［名前ボックス］", kind: "button" };
    if (/フィルハンドル|オートフィル/.test(p) && !/タブ/.test(p) && !/候補から/.test(p)) {
      return { label: "［コピー］", kind: "button" };
    }
    if (skip.test(p) && !/タブ/.test(p) && !/候補から/.test(p) && !/関数の挿入/.test(p)) continue;
    const tab = TABS.find((t) => t.re.test(p) && /タブ/.test(p));
    if (tab) return { label: tab.label, kind: "tab" };
    const fn = p.match(/候補から\s*[\[\［]([A-Z]+)[\]］]/);
    if (fn) return { label: "［" + fn[1] + "］", kind: "button" };
    if (/フィルターボタン/.test(p)) return { label: "［フィルター］", kind: "button" };
    const btn = p.match(/[\[［]([^\]］]+)[\]］]\s*(をクリック|のチェック|ボタン)/);
    if (btn && !/シート/.test(btn[1]) && btn[1] !== "=" && btn[1] !== "Enter" && btn[1] !== "F4") {
      const name = btn[1].replace(/\s+/g, "").replace(/･･･/g, "");
      if (name.length > 1 && name.length < 24) return { label: "［" + name + "］", kind: "button" };
    }
  }
  for (const p of parts) {
    if (skip.test(p) && !/タブ/.test(p)) continue;
    const m = p.match(/[\[［]([^\]］]+)[\]］]/);
    if (m && !/シート/.test(m[1]) && m[1].length > 1 && m[1] !== "=" && m[1] !== "Enter") {
      const name = m[1].replace(/\s+/g, "");
      if (/タブ/.test(p)) return { label: "［" + name.replace(/タブ/, "") + "］タブ", kind: "tab" };
      return { label: "［" + name + "］", kind: "button" };
    }
  }
  return { label: "［ホーム］タブ", kind: "tab", fallback: true };
}

const TAB_POOL = [
  "［ホーム］タブ", "［挿入］タブ", "［ページレイアウト］タブ", "［数式］タブ",
  "［データ］タブ", "［校閲］タブ", "［表示］タブ", "［ファイル］タブ",
  "［テーブルデザイン］タブ", "［グラフのデザイン］タブ", "［書式］タブ"
];
const BTN_POOL = [
  "［縞模様（行）］", "［セルのスタイル］", "［余白］", "［印刷の向き］",
  "［条件付き書式］", "［ウィンドウ枠の固定］", "［リンク］", "［並べ替え］",
  "［数式の表示］", "［ヘッダーとフッター］", "［関数の挿入］", "［コピー］",
  "［削除］", "［メモ］"
];

const CATS = {
  1: "テーブル", 2: "セルの書式", 3: "印刷設定", 4: "グラフ作成", 5: "グラフ編集",
  6: "ブックの設定", 7: "データ分析", 8: "関数", 9: "名前と文字列", 10: "関数"
};

function pickWrong(correct, kind, seed) {
  const pool = kind === "tab" ? TAB_POOL : BTN_POOL;
  const filtered = pool.filter((x) => x !== correct);
  const src = filtered.length ? filtered : TAB_POOL.filter((x) => x !== correct);
  return src[seed % src.length];
}

const questions = [];
const unmatched = [];
json.projects.forEach((p) => {
  p.tasks.forEach((t) => {
    const nq = norm(t.description);
    let best = null;
    let bestScore = 0;
    for (const c of csvItems) {
      const nc = norm(c.q);
      if (nq === nc || nq.includes(nc) || nc.includes(nq)) {
        best = c;
        bestScore = 100;
        break;
      }
      let s = 0;
      const len = Math.min(nq.length, nc.length);
      const startNq = nq.indexOf("シート");
      const startNc = nc.indexOf("シート");
      const a = startNq >= 0 ? nq.slice(startNq) : nq;
      const b = startNc >= 0 ? nc.slice(startNc) : nc;
      const n = Math.min(a.length, b.length, 50);
      for (let i = 0; i < n; i++) if (a[i] === b[i]) s++;
      if (s > bestScore) {
        bestScore = s;
        best = c;
      }
    }
    if (bestScore < 18) unmatched.push({ id: p.projectId + "-" + t.taskId, score: bestScore });
    const ex = extractFirst(best ? best.steps : "");
    questions.push({
      id: p.projectId + "-" + t.taskId,
      cat: CATS[p.projectId],
      q: t.description.replace(/\n/g, " "),
      correct: ex.label,
      wrong: pickWrong(ex.label, ex.kind, p.projectId * 10 + t.taskId),
      fallback: !!ex.fallback,
      score: bestScore
    });
  });
});

console.log("count", questions.length);
console.log("unmatched", JSON.stringify(unmatched));
questions.forEach((x) => {
  console.log(x.id + "\t" + x.correct + "\t" + x.wrong + "\t" + x.score + (x.fallback ? "\tFALLBACK" : ""));
});

const out = questions.map((x) => ({
  id: x.id,
  cat: x.cat,
  q: x.q,
  correct: x.correct,
  wrong: x.wrong
}));
const js = "// Excel 演習タブ（プロジェクト1〜10）の 61 問。correct は解答手順の最初のタブ/ボタン。\nconst QUESTIONS = " +
  JSON.stringify(out, null, 2) + ";\n";
fs.writeFileSync(path.join(__dirname, "questions.js"), js, "utf8");
console.log("wrote questions.js", out.length);
