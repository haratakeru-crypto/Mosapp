"""Generate Copilot MD from variant problem JSON."""
import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]


def task_params_table(s: dict, pid: int) -> str:
    rows = []
    for i in range(1, 20):
        key = f"task{pid}_{i}"
        if key not in s.get("taskParams", {}):
            break
        p = s["taskParams"][key]
        sheet = p.get("sheet", "—")
        parts = [f"{k}={v}" for k, v in p.items() if k != "sheet"]
        rows.append(f"| {pid}-{i} | {sheet} | {', '.join(parts)} |")
    header = "| taskId | sheet | パラメータ |"
    sep = "|--------|-------|------------|"
    return "\n".join([header, sep] + rows)


def layout_table(s: dict) -> str:
    ly = s["layout"]
    lines = [
        "| 項目 | 値 |",
        "|------|-----|",
        f"| 表開始セル | {ly.get('tableStart')} |",
        f"| ヘッダー行 | {ly.get('headerRow')}行目 |",
        f"| 表形状 | {ly.get('tableShape')} |",
        f"| シート数 | {ly.get('sheetCount')}枚 |",
    ]
    if ly.get("colWidths"):
        lines.append(f"| 列幅 | {', '.join(str(w) for w in ly['colWidths'])} |")
    rh = ly.get("rowHeights")
    if rh:
        lines.append(
            f"| 行高 | タイトル{rh.get('title')} / ヘッダー{rh.get('header')} / "
            f"データ{rh.get('data')} / 集計{rh.get('summary')} |"
        )
    lines += [
        f"| 視覚要素 | {ly.get('visual')} |",
        f"| 書式 | {ly.get('format')} |",
        f"| タブ色 | {ly.get('tabColor')} |",
        f"| 印刷 | {ly.get('print')} |",
    ]
    return "\n".join(lines)


def project5_dispersion_table(data: dict) -> list[str]:
    """タスク5-3 / 5-5 のグラフ要素分散表（Project5 専用）。"""
    sets = data["sets"]
    rows_53 = []
    rows_55 = []
    for s in sets:
        p3 = s["taskParams"].get("task5_3", {})
        p5 = s["taskParams"].get("task5_5", {})
        t3 = next((t for t in s["tasks"] if t.startswith("タスク5-3")), "")
        t5 = next((t for t in s["tasks"] if t.startswith("タスク5-5")), "")
        elem3 = p3.get("chartElementUi") or p3.get("chartTitle") or "—"
        elem5 = p5.get("chartElementUi") or "—"
        rows_53.append(f"| {s['setNo']} {s['theme']} | {p3.get('sheet', '—')} | {elem3} | {t3.split('の', 1)[-1] if 'の' in t3 else t3} |")
        rows_55.append(f"| {s['setNo']} {s['theme']} | {p5.get('sheet', '—')} | {elem5} | {t5.split('の', 1)[-1] if 'の' in t5 else t5} |")
    return [
        "**タスク5-3 のグラフ要素（セットごとに異なる）:**",
        "",
        "| セット | シート | 要素 | 操作概要 |",
        "|--------|--------|------|----------|",
        *rows_53,
        "",
        "**タスク5-5 のグラフ要素（セットごとに異なる）:**",
        "",
        "| セット | シート | 要素 | 操作概要 |",
        "|--------|--------|------|----------|",
        *rows_55,
        "",
    ]


def generate(pid: int) -> str:
    data = json.loads((BASE / "類題Json" / f"MOS_類題_project{pid}_配置別_5セット_問題文.json").read_text(encoding="utf-8"))
    forbidden = data.get("variantRules", {}).get("forbiddenSheetNames", [])
    task_count = len(data["sets"][0]["tasks"])
    lines = [
        f"# Excel 類題作成 — Copilot 用プロンプト（Project{pid} 配置別）",
        "",
        "Microsoft Copilot に **コピペして使う** プロンプト集です。",
        f"project{pid} の類題ブック（5セット×{task_count}タスク）を作成します。",
        "",
        "**問題文ファイル（完成初稿）:**",
        f"- [`task/MOS_類題_project{pid}_配置別_5セット_問題文.md`](MOS_類題_project{pid}_配置別_5セット_問題文.md)",
        f"- [`task/類題Json/MOS_類題_project{pid}_配置別_5セット_問題文.json`](類題Json/MOS_類題_project{pid}_配置別_5セット_問題文.json)",
        "",
        f"**禁止シート名（教材流用不可）:** {' '.join(f'`{x}`' for x in forbidden)}",
        "",
    ]
    if pid == 5:
        ui = data.get("excelUiNames", {})
        lines += [
            "**グラフ要素の UI パス（ビルド2508）:**",
            "",
            f"- クイックレイアウト: {ui.get('chartLayout', '—')}",
            f"- グラフスタイル: {ui.get('chartStyle', '—')}",
            f"- 色の変更: {ui.get('chartColor', '—')}",
            f"- グラフタイトル: {ui.get('chartTitle', '—')}",
            f"- データテーブル: {ui.get('dataTable', '—')}",
            f"- グリッド線: {ui.get('gridlines', '—')}",
            f"- 目盛り: {ui.get('axisTicks', '—')}",
            f"- 凡例: {ui.get('legend', '—')}",
            f"- 軸のタイトル: {ui.get('axisTitle', '—')}",
            "",
            *project5_dispersion_table(data),
        ]
    lines += [
        "---",
        "",
        "## 0. 使い方（3セッション × 1セット）",
        "",
        "- 問題文はセッション3まで書かない",
        "- 操作種別を変更しない（5-1=レイアウト+タイトル、5-2=スタイル+配色、5-3〜6=グラフ要素）" if pid == 5 else "- 操作種別を変更しない",
        f"- ブック名: `MOS_類題_project{pid}_配置別_セット{{N}}_{{テーマ}}.xlsx`",
        "",
        "セッション2/3 は Project4 Copilot MD §0 の共通テンプレを流用し、下記 taskParams 表を貼る。",
        "",
    ]
    for s in data["sets"]:
        n = s["setNo"]
        lines += [
            f"## §{n} セット{n} — {s['theme']}",
            "",
            f"**ブック名:** `{s['workbook']}`",
            "",
            "### セッション1 — レイアウト仕様",
            "",
            layout_table(s),
            "",
            f"**シート一覧:** {', '.join(s['layout']['sheets'])}",
            "",
            "### セッション2 / セッション3 — 設計値",
            "",
            task_params_table(s, pid),
            "",
            "---",
            "",
        ]
    lines += ["## 7. 完成問題文（初稿）", ""]
    for s in data["sets"]:
        lines += [f"### {s['workbook']}", "", f"{s['problemStatement']}  ", "問題）  "]
        for t in s["tasks"]:
            lines.append(t)
        lines.append("")
    lines += [
        "## 改訂履歴",
        "",
        "| 日付 | 内容 |",
        "|------|------|",
        f"| 2026-06-25 | Project{pid} 配置別5セット初版 |",
    ]
    if pid == 5:
        lines.insert(-1, "| 2026-06-25 | タスク5-3/5-5 をグラフ要素（タイトル・データテーブル・凡例・目盛り・軸ラベル等）に分散 |")
    lines += [
        "",
    ]
    return "\n".join(lines)


if __name__ == "__main__":
    for p in sys.argv[1:] or range(5, 11):
        pid = int(p)
        out = BASE / f"Excel_類題_Copilot_project{pid}_配置別.md"
        out.write_text(generate(pid), encoding="utf-8")
        print(f"wrote {out}")
