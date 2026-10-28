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

function csvEscape(s) {
  return `"${String(s).replace(/"/g, '""')}"`;
}

function pickWrong(correct, kind, seed, tabs, btns) {
  const pool = (kind === "tab" ? tabs : btns).filter((x) => x !== correct);
  const src = pool.length ? pool : tabs.filter((x) => x !== correct);
  return src[seed % src.length];
}

function extractFirst(steps, tabs, skipRe) {
  const raw = (steps || "").replace(/\r\n/g, "\n").replace(/\n+/g, " ");
  const parts = raw.split("↓").map((s) => s.replace(/\s+/g, " ").trim()).filter(Boolean);
  for (const p of parts) {
    if (skipRe.test(p) && !/タブ/.test(p)) continue;
    const tab = tabs.find((t) => t.re.test(p) && /タブ/.test(p));
    if (tab) return { label: tab.label, kind: "tab", note: p };
    const btn = p.match(/[\[［]([^\]］]+)[\]］]\s*(をクリック|を選択|をポイント|ボタン)/);
    if (btn && !/スライド|シート/.test(btn[1]) && btn[1].length > 1 && btn[1].length < 28) {
      const name = btn[1].replace(/\s+/g, "");
      if (!/タブ$/.test(name)) return { label: "［" + name + "］", kind: "button", note: p };
    }
  }
  for (const p of parts) {
    if (skipRe.test(p) && !/タブ/.test(p)) continue;
    const tab = tabs.find((t) => t.re.test(p) && /タブ/.test(p));
    if (tab) return { label: tab.label, kind: "tab", note: p };
    const m = p.match(/[\[［]([^\]］]+)[\]］]/);
    if (m && m[1].length > 1 && !/スライド|シート/.test(m[1])) {
      const name = m[1].replace(/\s+/g, "");
      if (/タブ/.test(p) || /タブ$/.test(name)) {
        return { label: "［" + name.replace(/タブ/, "") + "］タブ", kind: "tab", note: p };
      }
      return { label: "［" + name + "］", kind: "button", note: p };
    }
  }
  return { label: "［ホーム］タブ", kind: "tab", fallback: true, note: parts[0] || "" };
}

function bestMatch(nq, items) {
  let best = null;
  let bestScore = 0;
  for (const c of items) {
    const nc = norm(c.q);
    if (!nc) continue;
    if (nq === nc || nq.includes(nc) || nc.includes(nq)) return { item: c, score: 100 };
    let s = 0;
    const a = nq.length >= nc.length ? nq : nc;
    const b = nq.length >= nc.length ? nc : nq;
    if (a.includes(b.slice(0, Math.min(18, b.length)))) s = 40;
    const n = Math.min(nq.length, nc.length, 40);
    for (let i = 0; i < n; i++) if (nq[i] === nc[i]) s++;
    if (s > bestScore) {
      bestScore = s;
      best = c;
    }
  }
  return { item: best, score: bestScore };
}

function writeCsv(outPath, rows) {
  const lines = ["id,cat,q,correct,wrong,steps_note"];
  rows.forEach((r) => {
    lines.push([r.id, r.cat, r.q, r.correct, r.wrong, r.steps_note].map(csvEscape).join(","));
  });
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, "\uFEFF" + lines.join("\r\n") + "\r\n", "utf8");
}

const WORD_TABS = [
  { re: /参考資料/, label: "［参考資料］タブ" },
  { re: /差し込み文書/, label: "［差し込み文書］タブ" },
  { re: /テーブル.?レイアウト|表の\s*［?レイアウト/, label: "［テーブルレイアウト］タブ" },
  { re: /図の形式/, label: "［図の形式］タブ" },
  { re: /図形の書式/, label: "［図形の書式］タブ" },
  { re: /Smart\s*Artのデザイン|スマートアートのデザイン/, label: "［SmartArtのデザイン］タブ" },
  { re: /デザイン/, label: "［デザイン］タブ" },
  { re: /レイアウト/, label: "［レイアウト］タブ" },
  { re: /ホーム/, label: "［ホーム］タブ" },
  { re: /挿入/, label: "［挿入］タブ" },
  { re: /校閲/, label: "［校閲］タブ" },
  { re: /表示/, label: "［表示］タブ" },
  { re: /ファイル/, label: "［ファイル］タブ" }
];
const WORD_TAB_POOL = WORD_TABS.map((t) => t.label);
const WORD_BTN_POOL = [
  "［箇条書き］", "［切り取り］", "［貼り付け］", "［余白］", "［区切り］",
  "［新しいコメント］", "［透かし］", "［画像］", "［表］", "［置換］",
  "［リンク］", "［目次］", "［返信］", "［1から再開］"
];
const WORD_SKIP = /を選択$|を選択↓|文字列を選択|段落を選択|先頭をクリック|カーソルを移動|図形を選択|テキストボックスを選択|画像をクリック|画像を選択|右クリック|表内にカーソル|空白の段落をクリック|後ろをクリック/;
const WORD_CATS = {
  1: "スタイルと書式", 2: "編集", 3: "ページ設定", 4: "デザインとコメント", 5: "図",
  6: "表", 7: "文書の管理", 8: "参考資料", 9: "表と校閲", 10: "検索と置換"
};
const WORD_HAND = {
  "2-4": { correct: "［ホーム］タブ", kind: "tab", note: "図形選択後に直接入力（タブなし。2択用に［ホーム］タブ）" },
  "4-2": { correct: "［返信］", kind: "button", note: "コメントの［返信］をクリック（タブなし）" },
  "4-3": { correct: "［その他のスレッド操作］", kind: "button", note: "コメントの［その他のスレッド操作］（タブなし）" },
  "4-5": { correct: "［デザイン］タブ", kind: "tab", note: "JSONは透かし「サンプル２」、手順CSVは「下書き1」。いずれも［デザイン］タブ" },
  "4-6": { correct: "［デザイン］タブ", kind: "tab", note: "JSONはオリーブ罫線、手順CSVは青。いずれも［デザイン］タブ" },
  "9-1": { correct: "［1から再開］", kind: "button", note: "段落番号を右クリック→［1から再開］（タブなし）" },
  "10-3": { correct: "［箇条書き］", kind: "button", note: "段落選択後 →［箇条書き］の下矢印（タブなし）" }
};

const PP_TABS = [
  { re: /画面切り替え|切り替え/, label: "［画面切り替え］タブ" },
  { re: /アニメーション/, label: "［アニメーション］タブ" },
  { re: /スライドショー/, label: "［スライドショー］タブ" },
  { re: /グラフィックス形式|グラフィックの形式/, label: "［グラフィックス形式］タブ" },
  { re: /3D\s*モデル/, label: "［3Dモデル］タブ" },
  { re: /テーブル.?デザイン|表のデザイン/, label: "［テーブルデザイン］タブ" },
  { re: /グラフのデザイン/, label: "［グラフのデザイン］タブ" },
  { re: /図の形式/, label: "［図の形式］タブ" },
  { re: /図形の書式/, label: "［図形の書式］タブ" },
  { re: /Smart\s*Artのデザイン|スマートアートのデザイン/, label: "［SmartArtのデザイン］タブ" },
  { re: /再生/, label: "［再生］タブ" },
  { re: /デザイン/, label: "［デザイン］タブ" },
  { re: /ホーム/, label: "［ホーム］タブ" },
  { re: /挿入/, label: "［挿入］タブ" },
  { re: /校閲/, label: "［校閲］タブ" },
  { re: /表示/, label: "［表示］タブ" },
  { re: /ファイル/, label: "［ファイル］タブ" }
];
const PP_TAB_POOL = PP_TABS.map((t) => t.label);
const PP_BTN_POOL = [
  "［新しいスライド］", "［レイアウト］", "［非表示スライド］", "［ズーム］",
  "［図］", "［表］", "［SmartArt］", "［画面切り替え］", "［アニメーション］",
  "［スライドショー］", "［リンク］", "［図形］", "［メディアの挿入］",
  "［グラフの挿入］", "［セクション名の変更］", "［SmartArtグラフィックの挿入］"
];
const PP_SKIP = /スライド\s*\d|スライドを選択|スライド\d+|プレースホルダーを選択|プレースホルダー以外|オブジェクトを選択|図形を選択|画像を選択|ビデオを選択|アイコンを選択|3D\s*モデルを選択|右クリック|選択を解除/;
const PP_CATS = {
  1: "スライド操作", 2: "画面切り替えとアニメーション", 3: "図と図形", 4: "表とSmartArt",
  5: "デザイン", 6: "画面切り替え", 7: "配布と校閲", 8: "スライドショー", 9: "図の編集", 10: "アニメーション"
};

function build(jsonPath, csvPath, qCol, stepsCol, cats, tabs, tabPool, btnPool, skipRe, hand, startRow) {
  const json = JSON.parse(fs.readFileSync(jsonPath, "utf8").replace(/^\uFEFF/, ""));
  const csvRows = parseCsv(fs.readFileSync(csvPath, "utf8")).slice(startRow);
  const items = csvRows.filter((r) => r.length > stepsCol).map((r) => ({
    q: r[qCol],
    steps: r[stepsCol]
  }));
  const out = [];
  const warnings = [];
  json.projects.forEach((p) => {
    p.tasks.forEach((t) => {
      const id = `${p.projectId}-${t.taskId}`;
      const q = (t.description || "").replace(/\r\n/g, "\n").replace(/\n/g, " ");
      const nq = norm(t.description);
      const m = bestMatch(nq, items);
      let ex = extractFirst(m.item ? m.item.steps : "", tabs, skipRe);
      if (hand[id]) {
        ex = { label: hand[id].correct, kind: hand[id].kind, note: hand[id].note };
      }
      if (m.score < 18 || ex.fallback) warnings.push(`${id}\tscore=${m.score}\t${ex.label}${ex.fallback ? "\tFALLBACK" : ""}`);
      const wrong = pickWrong(ex.label, ex.kind, p.projectId * 10 + t.taskId, tabPool, btnPool);
      const note = (ex.note || "").replace(/\s+/g, " ").slice(0, 80);
      out.push({
        id,
        cat: cats[p.projectId] || "",
        q,
        correct: ex.label,
        wrong,
        steps_note: (hand[id] ? hand[id].note : note) + (m.score < 100 ? `（突合せscore ${m.score}）` : "")
      });
    });
  });
  return { out, warnings };
}

const word = build(
  path.join(root, "MOS Word app/References/JSON/MOS模擬アプリ問題文一覧_Word.json"),
  path.join(root, "MOS Word app/Reference/CSV/MOS模擬アプリ正誤判定表251120_.csv"),
  2, 3, WORD_CATS, WORD_TABS, WORD_TAB_POOL, WORD_BTN_POOL, WORD_SKIP, WORD_HAND, 1
);
const PP_HAND = {
  "1-7": { correct: "［セクション名の変更］", kind: "button", note: "セクション名を右クリック→［セクション名の変更］（タブなし）" },
  "3-5": { correct: "［3Dモデル］タブ", kind: "tab", note: "3Dモデル選択後 →［3Dモデル］タブ→［ビュー］" },
  "4-1": { correct: "［ホーム］タブ", kind: "tab", note: "図形選択後に直接入力（タブなし。2択用に［ホーム］タブ）" },
  "5-7": { correct: "［グラフィックス形式］タブ", kind: "tab", note: "アイコン選択後 →［グラフィックス形式］タブ" },
  "9-1": { correct: "［メディアの挿入］", kind: "button", note: "スライド選択後 →［メディアの挿入］アイコン（タブなし）" },
  "9-2": { correct: "［再生］タブ", kind: "tab", note: "ビデオ選択後 →［再生］タブ→［ビデオのトリミング］" },
  "9-4": { correct: "［グラフの挿入］", kind: "button", note: "スライド選択後 →［グラフの挿入］アイコン（タブなし）" },
  "9-5": { correct: "［グラフのデザイン］タブ", kind: "tab", note: "グラフ選択後 →［グラフのデザイン］タブ" }
};

const ppt = build(
  path.join(root, "Mos PowerPoint Mogi App/References/JSON/MOS模擬アプリ問題文一覧_PowerPoint.json"),
  path.join(root, "../task/PP修正版問題文一覧.csv"),
  3, 4, PP_CATS, PP_TABS, PP_TAB_POOL, PP_BTN_POOL, PP_SKIP, PP_HAND, 1
);

writeCsv(path.join(__dirname, "word-3sec-snipe/choices-review.csv"), word.out);
writeCsv(path.join(__dirname, "ppt-3sec-snipe/choices-review.csv"), ppt.out);
console.log("word", word.out.length, "warn", word.warnings.length);
word.warnings.forEach((w) => console.log("W", w));
console.log("ppt", ppt.out.length, "warn", ppt.warnings.length);
ppt.warnings.forEach((w) => console.log("P", w));
console.log(word.out.map((x) => x.id + "\t" + x.correct + "\t" + x.wrong).join("\n"));
console.log("---PPT---");
console.log(ppt.out.map((x) => x.id + "\t" + x.correct + "\t" + x.wrong).join("\n"));
