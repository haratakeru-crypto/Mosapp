"""Generate Copilot MD from PowerPoint variant problem JSON."""

from __future__ import annotations

import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
TOOLS = Path(__file__).resolve().parent
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
from ppt_copilot_content import content_md_path, load_content_for_project
from ppt_slide_layout_hints import render_exam_layout_hints
from ppt_slide_outline import render_slide_content_block, render_slide_map_md

FIXED_CONTENT_PROJECTS = frozenset({4, 5, 6, 7, 8, 9, 10})


def usage_section_fixed_content(pid: int) -> list[str]:
    content_file = content_md_path(pid).name
    lines = [
        "## 0. 使い方 — 固定コンテンツ転記（Project4〜10）",
        "",
        f"**本文・箇条書き・スピーカーノートの正本は `{content_file}` です。**",
        "Copilot は創作せず、各セットの「貼り付け用スライドデータ」を一字一句転記してください。",
        "",
        "【推奨: 2段階生成】",
        "- 第1段階: 本文・表データまで作成（メディア挿入・マスター操作は未実施）",
        "- 第2段階: 完成pptxを添付し、試験用レイアウト仕様どおりにメディア・配置のみ修正",
        "",
        "【セッション1で禁止】",
        "- 箇条書き・段落・ノートの文言変更・追加・削除",
        "- 「項目A」「ダミー」「...」などのプレースホルダ",
        "- タイトル・図形ラベル・taskParams 文字列の変更",
        "- 受験者操作の事前完了（位置揃え・トリミング・塗りつぶし・グループ化・マスター変更 等）",
        "- アイコン／SmartArt／汎用プレースホルダで画像を代用すること",
        "",
        "【セッション1で必須】",
        "- 全スライドのコンテンツ領域に貼り付け用データを転記",
        "- 「試験用レイアウト仕様」の図形・画像配置を厳守",
        "- 保存前セルフチェック: 空スライドなし / 禁止語なし / レイアウト仕様どおり",
        "",
    ]
    if pid in (9, 10):
        lines += [
            f"【Project{pid} 追加注意】",
            "- P9: ビデオ・オーディオは「挿入→メディアのこのデバイス」。Copilot生成イラスト禁止",
            "- P10: 通常スライド本文のみ Copilot 作成。マスター操作は受験者用のため初期状態はマスター未変更",
            "",
        ]
    lines.append("---")
    lines.append("")
    return lines


def variation_table(data: dict) -> str:
    pid = data["projectId"]
    rows = []
    for s in data["sets"]:
        tp = s["taskParams"]
        if pid == 1:
            row = (
                f"| {s['setNo']} {s['theme']} | {tp['task1_1']['layoutUi']} | "
                f"{tp['task1_5']['spacingPt']}pt | {tp['task1_8']['zoomTitle']} |"
            )
        elif pid == 2:
            row = (
                f"| {s['setNo']} {s['theme']} | {tp['task2_1']['transition']} | "
                f"{tp['task2_2']['durationSec']}秒 | {tp['task2_4']['autoAdvanceSec']}秒 |"
            )
        elif pid == 3:
            row = (
                f"| {s['setNo']} {s['theme']} | {tp['task3_1']['smartArtType']} | "
                f"{tp['task3_2']['smartArtColor']} | {tp['task3_4']['modelName']} |"
            )
        elif pid == 4:
            row = (
                f"| {s['setNo']} {s['theme']} | {tp['task4_1']['inputShapeType']} | "
                f"{tp['task4_7']['textBoxShapeType']} | {tp['task4_1']['shapeText']} |"
            )
        elif pid == 5:
            row = (
                f"| {s['setNo']} {s['theme']} | {tp['task5_1']['alignShapeType']} | "
                f"{tp['task5_3']['sourceShapeType']}→{tp['task5_3']['targetShapeType']} | "
                f"{tp['task5_5']['gateShapeType']} |"
            )
        elif pid == 6:
            row = (
                f"| {s['setNo']} {s['theme']} | {tp['task6_4']['customShowName']} | "
                f"{tp['task6_5']['outlineCopies']}部 | {tp['task6_7']['handoutCopies']}部 |"
            )
        elif pid == 7:
            row = (
                f"| {s['setNo']} {s['theme']} | {tp['task7_2']['hyperlinkUrl']} | "
                f"{tp['task7_4']['footerDomain']} | {tp['task7_5']['footerText']} |"
            )
        elif pid == 8:
            row = (
                f"| {s['setNo']} {s['theme']} | {tp['task8_1']['tableStyle']} | "
                f"{tp['task8_2']['backgroundColor']} | "
                f"{tp['task8_4']['widthCm']}×{tp['task8_4']['heightCm']}cm |"
            )
        elif pid == 9:
            row = (
                f"| {s['setNo']} {s['theme']} | {tp['task9_2']['trimStartSec']}-{tp['task9_2']['trimEndSec']}秒 | "
                f"{tp['task9_4']['categoryColumn']} | {tp['task9_1']['videoFile']} |"
            )
        elif pid == 10:
            row = (
                f"| {s['setNo']} {s['theme']} | {tp['task10_1']['masterTheme']} | "
                f"{tp['task10_8']['handoutFooter']} | {tp['task10_6']['customLayoutName']} |"
            )
        else:
            row = f"| {s['setNo']} {s['theme']} | — | — | — |"
        rows.append(row)
    headers = {
        1: "| セット | 1-1レイアウト | 1-5間隔 | 1-8ズームタイトル |",
        2: "| セット | 2-1切替 | 2-2秒 | 2-4自動秒 |",
        3: "| セット | 3-1SmartArt | 3-2色 | 3-4モデル |",
        4: "| セット | 4-1入力図形 | 4-7テキスト枠形状 | 4-1入力文字 |",
        5: "| セット | 5-1整列図形 | 5-3図形変更 | 5-5ゲート形状 |",
        6: "| セット | 6-4目的別名 | 6-5アウトライン部数 | 6-7配布部数 |",
        7: "| セット | 7-2URL | 7-4フッタードメイン | 7-5事例フッター |",
        8: "| セット | 8-1表スタイル | 8-2背景色 | 8-4サイズ(cm) |",
        9: "| セット | 9-2トリム秒 | 9-4グラフ列 | 9-1ビデオ |",
        10: "| セット | 10-1マスターテーマ | 10-8配布フッター | 10-6カスタムレイアウト |",
    }
    header = headers[pid]
    return "\n".join([f"**Project{pid} パラメータ一覧**", "", header, "|--------|" + "------|" * (header.count("|") - 2), *rows, ""])


def _fixed_content_lines(pid: int, set_no: int, content_by_set: dict) -> list[str]:
    slides = content_by_set.get(set_no)
    if not slides:
        return [
            "【貼り付け用スライドデータ — 全スライド必須転記】",
            f"  ※ Content MD にセット{set_no}の定義がありません。PowerPoint_類題_Copilot_Content_project{pid}.md を確認してください。",
            "",
        ]
    return [
        "【最重要】本文・箇条書き・ノートの創作は禁止。下記「貼り付け用スライドデータ」を一字一句転記すること。",
        "",
        "【貼り付け用スライドデータ — 全スライド必須転記】",
        *render_slide_content_block(slides),
        "【保存前セルフチェック — 本文】",
        "- 全スライドのコンテンツ領域に上記箇条書きが入っている",
        "- 禁止語句が含まれていない",
        "- 同一文のコピペがない",
        "",
    ]


def session1(s: dict, pid: int, forbidden: list[str], content_by_set: dict | None = None) -> str:
    ly = s["layout"]
    cb = s["contentBlocks"]
    forbidden_line = "、".join(forbidden[:6]) + " 等"
    slide_lines = render_slide_map_md(cb.get("slideMap", []), cb.get("preTaskState"))
    layout_lines = render_exam_layout_hints(pid, cb.get("slideMap", []), s.get("taskParams", {}))
    pre = cb.get("preTaskState", {})
    pre_lines = [f"- {k}: 未実施" for k, v in pre.items() if not v]
    content_by_set = content_by_set or {}
    body: list[str] = [
        f"## セット{s['setNo']}：{s['workbook']}",
        "",
        f"**テーマ:** {s['theme']} / **色:** {ly.get('colorTheme')} / **デザイン:** {ly.get('designTheme')}",
        "",
        "### プロンプト（セッション1：完全 pptx 作成）",
        "",
        "```text",
        "あなたは MOS PowerPoint 365 演習用の .pptx を作成するアシスタントです。",
        "以下の仕様どおりに **完全なプレゼンテーション** を作成し、ファイルを保存してください。",
        "骨組みのみ・空欄プレースホルダ・「ダミー」文字は禁止です。",
        "",
        f"【禁止語句】{forbidden_line}",
        "",
        f"【導入】{s['problemStatement']}",
        "",
        "【デザイン】",
        f"- スライドサイズ: {ly.get('slideSize')}",
        f"- デザイン テーマ: {ly.get('designTheme')} / バリアント: {ly.get('variant')}",
        f"- 色のテーマ: {ly.get('colorTheme')}",
        f"- フォント: {ly.get('fontTheme')}",
        "",
    ]
    if pid in FIXED_CONTENT_PROJECTS:
        body.extend(_fixed_content_lines(pid, s["setNo"], content_by_set))
        if layout_lines:
            body.extend(layout_lines)
            layout_checks = [
                "【保存前セルフチェック — レイアウト】",
                "- 挿入→画像／図形／アイコンを使用（汎用プレースホルダ図形禁止）",
                "- 受験者操作（揃え・トリミング・塗りつぶし等）は未実施",
                "- 図形ラベル・タイトルが slideMap と一致",
            ]
            if pid == 4:
                layout_checks.insert(
                    2,
                    "- スライド5: 画像2枚は横並び、上端ずれ・右はみ出しが初期状態のまま",
                )
            if pid == 9:
                layout_checks.insert(
                    2,
                    "- P9: メディアは実ファイル挿入。グラフ・表データは未加工の初期状態",
                )
            if pid == 10:
                layout_checks.insert(
                    2,
                    "- P10: スライドマスターは未変更のまま（本文スライドのみ作成）",
                )
            layout_checks.append("")
            body.extend(layout_checks)
    body.extend(
        [
            "【スライド構成 — 図形・オブジェクト一覧（参照）】",
            *[line.replace("- ", "  - ") if line.startswith("-") else line for line in slide_lines],
            "",
            "【受験者操作前 — 以下は未実施のまま】",
            *pre_lines,
            "",
            "【問題文（参考・操作は受験者が実施）】",
            *[t.replace("タスク", "  - タスク") for t in s["tasks"]],
            "",
            f"ファイル名: {s['workbook']}",
            "```",
            "",
        ]
    )
    return "\n".join(body)


def batch_prompt(pid: int, data: dict) -> str:
    sets = data["sets"]
    names = "\n".join(f"  - {s['workbook']}" for s in sets)
    extra: list[str] = []
    if pid in FIXED_CONTENT_PROJECTS:
        extra = [
            "各セットの「貼り付け用スライドデータ」は一字一句転記し、本文の創作は禁止です。",
            "「試験用レイアウト仕様」の図形・画像配置を厳守してください（P4スライド5は横並び2枚・上端ずれ・右はみ出し）。",
            "図形は slideMap の shapes 配列の insertName（挿入→図形）で配置し、type と問題文の図形呼称を一致させてください。",
            "P9/P10 は第1段階で本文・表まで作成し、メディア・マスターは手動後工程でも可。",
        ]
    else:
        extra = [
            "図形は slideMap の shapes 配列の insertName（挿入→図形）で配置し、type と問題文の図形呼称を一致させてください。",
        ]
    return "\n".join(
        [
            "## §0b 5セット一括作成プロンプト",
            "",
            "```text",
            f"PowerPoint Project{pid} の類題 pptx を5セット分まとめて作成してください。",
            "各セットは独立したファイルとして保存します。",
            "",
            "作成するファイル:",
            names,
            "",
            "各セットについて、本ドキュメントの「セットN」セッション1プロンプトの仕様を厳守してください。",
            "スライド番号・タイトル文字列・オブジェクト名はパラメータ表と完全一致させてください。",
            *extra,
            "受験者が行う操作（非表示・ズーム・SmartArt色等）は事前に完了させないでください。",
            "```",
            "",
        ]
    )


def generate(pid: int) -> str:
    json_path = BASE / "類題Json" / f"MOS_PowerPoint類題_project{pid}_配置別_5セット_問題文.json"
    data = json.loads(json_path.read_text(encoding="utf-8"))
    forbidden = data.get("variantRules", {}).get("forbiddenTexts", [])
    content_by_set = load_content_for_project(pid) if pid in FIXED_CONTENT_PROJECTS else {}
    parts = [
        f"# PowerPoint 類題 Copilot 用 — Project{pid} 配置別（5セット）",
        "",
        "MOS本番相当の **完全な pptx** を作成するためのプロンプト集です。",
        "",
    ]
    if pid in FIXED_CONTENT_PROJECTS:
        parts.extend(usage_section_fixed_content(pid))
    parts.extend(
        [
            variation_table(data),
            batch_prompt(pid, data),
            "---",
            "",
        ]
    )
    for s in data["sets"]:
        parts.append(session1(s, pid, forbidden, content_by_set))
    return "\n".join(parts)


def main() -> None:
    pids = [int(a) for a in sys.argv[1:]] if len(sys.argv) > 1 else [1]
    for pid in pids:
        out = BASE / f"PowerPoint_類題_Copilot_project{pid}_配置別.md"
        out.write_text(generate(pid), encoding="utf-8")
        print(f"wrote {out}")


if __name__ == "__main__":
    main()
