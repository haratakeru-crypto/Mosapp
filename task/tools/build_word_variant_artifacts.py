"""Generate MD / app JSON from Word variant problem JSON. Validate consistency."""
import argparse
import json
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
TOOLS = Path(__file__).resolve().parent
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
from word_document_outline import render_outline_md

PROJECT_CONFIG = {
    1: {
        "task_count": 5,
        "forbidden": {
            "CSR活動のメリット",
            "社員のコンプライアンス意識の確立",
            "CSR活動の光と影",
            "会社の社会的責任",
            "ゼミ",
            "第１節　はじめに",
            "第２節 会社の社会的責任の定義と問題点",
        },
    },
    2: {
        "task_count": 5,
        "forbidden": {
            "朗読会",
            "朗読を楽しみましょう",
            "青空文庫のURLはコチラ↓",
            "方法",
            "ご参加お待ちしております",
        },
    },
    3: {
        "task_count": 6,
        "forbidden": {
            "参考文献一覧",
            "茨城県天心記念五浦美術館",
            "美術館",
            "五浦",
        },
    },
    4: {
        "task_count": 7,
        "forbidden": {
            "1.生活の中でできるエコ活動",
            "その他",
            "エコと節約",
            "前田先生に最終確認",
            "最新の情報を確認",
            "サンプル２",
            "エコ活動",
        },
    },
    5: {
        "task_count": 8,
        "forbidden": {
            "5月21日より5日間",
            "TOEICテスト対策セミナー",
            "TOEIC",
            "セミナー案内",
        },
    },
    6: {
        "task_count": 7,
        "forbidden": {
            "月別催事内容",
            "実施月",
            "富士の水だより",
            "北海道",
            "うまいもの市場",
            "最高売上",
            "平均売上",
            "合計売上",
        },
        "content_required": [
            "documentOutline",
            "introParagraphs",
            "task6_1_tabRows",
            "task6_4_table",
            "task6_7_trailingBlankPage",
        ],
    },
    7: {
        "task_count": 8,
        "forbidden": {
            "朗読会",
            "ラビット出版",
            "インテグラル",
            "abc",
        },
        "content_required": ["documentOutline", "introParagraphs", "bodyParagraphs"],
    },
    8: {
        "task_count": 7,
        "forbidden": {
            "サイバーセキュリティ",
            "マルウェア",
            "参考文献一覧",
            "機密性",
            "完全性",
            "可用性",
            "TOP",
        },
        "content_required": [
            "documentOutline",
            "subtitle",
            "bodyUnderHeadings",
            "bibliographyEntries",
            "linkShape",
        ],
    },
    9: {
        "task_count": 6,
        "forbidden": {
            "ゴールデンウィーク",
            "月別催事内容",
            "小豆島",
            "654",
        },
        "content_required": [
            "documentOutline",
            "numberedSection",
            "task9_2_table",
            "task9_3_table",
            "trackChangesSamples",
        ],
    },
    10: {
        "task_count": 5,
        "forbidden": {
            "ウイルス",
            "コンピュータウイルス",
            "最低限意識してほしいこと",
            "その他意識してほしいこと",
            "事例",
            "PCアイコン",
        },
        "content_required": [
            "documentOutline",
            "bodyWithSearchTerms",
            "section10_2_paragraphs",
            "section10_3_paragraphs",
            "lastTwoLines",
        ],
    },
}


def json_path(project_id: int) -> Path:
    return BASE / "類題Json" / f"MOS_Word類題_project{project_id}_配置別_5セット_問題文.json"


def md_path(project_id: int) -> Path:
    return BASE / f"MOS_Word類題_project{project_id}_配置別_5セット_問題文.md"


def app_json_path(project_id: int) -> Path:
    return BASE / "類題Json" / f"MOS_Word類題_project{project_id}_配置別_5セット_問題文_アプリ用.json"


def _md_table(headers: list, rows: list) -> list[str]:
    lines = ["| " + " | ".join(headers) + " |", "| " + " | ".join(["---"] * len(headers)) + " |"]
    for row in rows:
        lines.append("| " + " | ".join(str(c) for c in row) + " |")
    return lines


def _p6_preop_state(s: dict) -> list[str]:
    p1 = s["taskParams"]["task6_1"]
    p2 = s["taskParams"]["task6_2"]
    p3 = s["taskParams"]["task6_3"]
    p5 = s["taskParams"]["task6_5"]
    p7 = s["taskParams"]["task6_7"]
    page = s["contentBlocks"].get("task6_7_trailingBlankPage", p7["page"])
    p5_done = (
        f"列幅 {', '.join(p5['columnWidthsCm'])}cm・行高 {p5['rowHeightCm']}cm は**未設定**"
        if p5.get("operation") == "setColumnWidthsAndRowHeight"
        else "列幅均等化は**未実施**（列幅は不均等のまま）"
    )
    return [
        "| タスク | 受験者操作前の docx 状態（必須） |",
        "|--------|------------------------------|",
        f"| 6-1 | 見出し「{p1['heading']}」下にタブ区切り **{p1['rowCount']}行×{p1['colCount']}列** のテキスト（**未表化**） |",
        f"| 6-2 | 6-1の表は存在するが、**{p2['splitAtRow']}行目からの分割は未実施** |",
        f"| 6-3 | 1つ目の表の **{p3['splitRow']}行{p3['splitCol']}列目** は未分割（受験者が{p3['splitInto']}列に分割） |",
        "| 6-4 | 第2表は完成データ入り・**1行目の文字効果は未設定**・列幅不均等 |",
        f"| 6-5 | {p5_done} |",
        "| 6-6 | タイトル行の繰り返しは**未設定** |",
        f"| 6-7 | **{page}ページ目末尾**に空段落のみ（{p7['rows']}行×{p7['cols']}列の表は**未作成**） |",
    ]


def _max_outline_level(sections: list, max_lv: int = 0) -> int:
    for sec in sections:
        max_lv = max(max_lv, sec.get("level", 0))
        for child in sec.get("children", []):
            max_lv = _max_outline_level([child], max_lv)
    return max_lv


def _append_document_outline_md(lines: list[str], cb: dict) -> None:
    outline = cb.get("documentOutline")
    if not outline:
        return
    lines.append("### 報告書の肉付け（見出し1〜4・本文）")
    lines.append("")
    lines.append("MOS本番に近い文書量にするため、以下の構成どおりに **見出しスタイル1〜4** と本文段落を配置してください。")
    lines.append("")
    lines.extend(render_outline_md(outline))
    lines.append("")


def _append_paste_blocks_md(lines: list[str], pid: int, s: dict) -> None:
    cb = s["contentBlocks"]
    ly = s["layout"]
    if pid == 6:
        p1 = s["taskParams"]["task6_1"]
        p4 = s["taskParams"]["task6_4"]
        p5 = s["taskParams"]["task6_5"]
        p7 = s["taskParams"]["task6_7"]
        tbl = cb["task6_4_table"]
        page = cb.get("task6_7_trailingBlankPage", p7["page"])
        lines.append("**導入段落:**")
        for p in cb.get("introParagraphs", []):
            lines.append(f"- {p}")
        lines.append("")
        lines.append(
            f"**6-1 用タブ区切りテキスト**（見出し「{p1['heading']}」の直下・**表に変換する前**）"
        )
        lines.append("")
        lines.append(f"行数={p1['rowCount']}、列数={p1['colCount']}。範囲は「{p1['startText']}」〜「{p1['endText']}」。")
        lines.append("")
        lines.append("```text")
        for row in cb["task6_1_tabRows"]:
            lines.append("\t".join(row))
        lines.append("```")
        lines.append("")
        lines.append(f"**6-4 用表「{p4['heading']}」**（列幅不均等・1行目文字効果なし）")
        lines.append("")
        lines.extend(_md_table(tbl["headers"], tbl["rows"]))
        lines.append("")
        if p5.get("operation") == "setColumnWidthsAndRowHeight":
            lines.append(
                f"※ 6-5 は列幅 **{', '.join(p5['columnWidthsCm'])} cm**（左から）・行高 **{p5['rowHeightCm']} cm** を"
                "受験者が設定する問題です。事前配置時は列幅・行高を**ばらばら**にしておくこと。"
            )
            if tbl.get("rowHeightsUnequal"):
                lines.append("（行の高さもセットごとに不均等にすること）")
            lines.append("")
        lines.append(
            f"**6-7 用:** {page}ページ目の**最後**に空の段落を1つ配置（"
            f"{p7['rows']}行×{p7['cols']}列の表はまだ作らない）。"
        )
        if p7.get("headerCells"):
            lines.append(
                f"1行目に入れるラベル（受験者が入力）: "
                + " / ".join(f"「{c}」" for c in p7["headerCells"])
            )
        lines.append("")
    elif pid == 7:
        lines += [
            "**互換モード docx:** 本文作成後 `.doc`（Word 97-2003）で保存し再オープン。会社名・ヘッダーは**未設定**。",
            "",
            "**導入段落:**",
        ]
        for p in cb["introParagraphs"]:
            lines.append(f"- {p}")
        lines.append("")
        lines.append("**本文段落:**")
        for p in cb["bodyParagraphs"]:
            lines.append(f"- {p}")
        lines.append("")
    elif pid == 8:
        ls = cb["linkShape"]
        lines += [
            f"**副題:** {cb['subtitle']}",
            "**目次:** 指定位置に未挿入（図形内・副題下の図形には入れない）",
            "",
            "**見出し別本文:**",
        ]
        for heading, paras in cb["bodyUnderHeadings"].items():
            lines.append(f"- 見出し「{heading}」:")
            for p in paras:
                lines.append(f"  - {p}")
        lines.append("")
        lines.append(f"**文献リスト（見出し「{s['taskParams']['task8_4']['heading']}」・特殊文字は**未挿入**）:**")
        for entry in cb["bibliographyEntries"]:
            lines.append(f"- {entry}")
        lines += [
            "",
            f"**リンク用図形:** {ls.get('note', '')}",
            f"  - テキスト「{ls['text']}」・配置: 見出し「{ls['nearHeading']}」下（リンク**未設定**）",
            "**最終ページ:** 空（SmartArt**未挿入**）・脚注**未挿入**",
            "",
        ]
    elif pid == 9:
        ns = cb["numberedSection"]
        t2 = cb["task9_2_table"]
        t3 = cb["task9_3_table"]
        p2 = s["taskParams"]["task9_2"]
        p3 = s["taskParams"]["task9_3"]
        p4 = s["taskParams"]["task9_4"]
        sep = p2.get("separatorLabel", "コンマ区切り")
        side = p3.get("marginSideLabel", "右")
        lines.append(f"**番号付き段落（見出し「{ns['heading']}」下）:**")
        for p in ns["paragraphs"]:
            lines.append(f"- {p}")
        lines.append("")
        lines.append(f"**表「{p2['heading']}」**（{sep}変換前）:")
        lines.extend(_md_table(t2["headers"], t2["rows"]))
        lines.append("")
        lines.append(
            f"**表「{p3['tableHeading']}」**"
            f"（{side}余白・「{p4['sortColumn']}」{p4.get('sortOrderLabel', '降順')}並べ替え**未実施**）:"
        )
        lines.extend(_md_table(t3["headers"], t3["rows"]))
        lines.append("")
        lines.append("**変更履歴（ON・未承認のまま）:**")
        for sample in cb["trackChangesSamples"]:
            lines.append(f"- [{sample['type']}] {sample['location']}: 「{sample['text']}」")
        lines.append("")
    elif pid == 10:
        p1 = s["taskParams"]["task10_1"]
        p3 = s["taskParams"]["task10_3"]
        p4 = s["taskParams"]["task10_4"]
        p5 = s["taskParams"]["task10_5"]
        lst = p5["lineSpacingType"]
        val = p5.get("lineSpacingValue")
        if lst in ("1行", "1.5行", "2行"):
            ls_label = lst
        elif lst == "倍数":
            ls_label = f"倍数{val}"
        else:
            ls_label = f"{lst}{val}pt"
        lines.append(f"**置換対象を含む本文（「{p1['searchText']}」が複数回）:**")
        for p in cb["bodyWithSearchTerms"]:
            lines.append(f"- {p}")
        lines.append("")
        lines.append(f"**見出し「{s['taskParams']['task10_2']['heading']}」下**（行頭文字・画像は**未設定**）:")
        for p in cb["section10_2_paragraphs"]:
            lines.append(f"- {p}")
        lines.append("")
        lines.append(
            f"**見出し「{p3['heading']}」下**"
            f"（{p3['bulletFont']} {p3['bulletCharCode']}は**未設定**）:"
        )
        for p in cb["section10_3_paragraphs"]:
            lines.append(f"- {p}")
        lines.append("")
        lines.append(f"**文書末尾2行**（用紙{p5['paperSize']}・行間{ls_label}は**未設定**）:")
        for p in cb["lastTwoLines"]:
            lines.append(f"- {p}")
        lines.append("")


PROJECT_OVERVIEW_OPS: dict[int, list[tuple[str, str]]] = {
    6: [
        ("6-1", "タブ区切りテキストを「文字列の幅に合わせた表」に変換"),
        ("6-2", "指定行から表を分割"),
        ("6-3", "指定セルを指定列数に分割"),
        ("6-4", "表1行目に文字の効果（塗りつぶし＋輪郭）"),
        ("6-5", "列幅をすべて同じにする **または** 列幅cm指定＋行の高さをすべて同じにする"),
        ("6-6", "表1行目を各ページのタイトル行に"),
        ("6-7", "指定ページ末尾に表を作成し1行目にラベルを入力"),
    ],
    7: [
        ("7-1", "互換モード解除"),
        ("7-2", "プロパティの会社名設定"),
        ("7-3", "ヘッダー挿入"),
        ("7-4", "テキストファイルとして保存"),
        ("7-5", "マクロ有効文書として保存（読み取りパスワード）"),
    ],
    8: [
        ("8-1", "指定位置に目次挿入（図形内・副題下以外）"),
        ("8-2", "脚注挿入"),
        ("8-3", "脚注開始番号変更（①以外）"),
        ("8-4", "見出し先頭に特殊文字（§以外）"),
        ("8-5", "SmartArt挿入（基本ベン図以外）"),
        ("8-6", "SmartArtの色変更（カラフル以外）"),
        ("8-7", "図形にリンク（最下部以外・見出し/URL/文末/文頭）"),
    ],
    9: [
        ("9-1", "段落番号の再開始（1以外）"),
        ("9-2", "表を文字列に変換（コンマ/タブ/段落区切り）"),
        ("9-3", "表セルの余白設定（右・4mm以外）"),
        ("9-4", "表の並べ替え（合計列・降順以外も）"),
        ("9-5", "変更履歴の承認・復元・取り消し"),
        ("9-6", "変更履歴のロック"),
    ],
    10: [
        ("10-1", "置換（置換前後はセットごとに独立）"),
        ("10-2", "行頭文字を画像に変更"),
        ("10-3", "行頭文字を記号フォントに変更（Webdings・120以外）"),
        ("10-4", "置換で書式変更（太字/色/サイズ）"),
        ("10-5", "用紙サイズ（B5以外）＋末尾2行の行間指定"),
    ],
}


def write_md_rich(data: dict) -> str:
    pid = data["projectId"]
    rules = data.get("variantRules", {})
    forbidden = rules.get("forbiddenParagraphTexts", [])
    lines = [
        f"# MOSスペシャリスト Word 類題 — Project{pid} 配置別（5セット）",
        "",
        f"元ファイル {data['sourceFile']} の操作をもとに、",
        "レイアウトを差別化した 5 セットの類題文書用問題文です。",
        "",
        "**関連ファイル:**",
        f"- 正本 JSON: `task/類題Json/MOS_Word類題_project{pid}_配置別_5セット_問題文.json`",
        f"- アプリ用 JSON: `task/類題Json/MOS_Word類題_project{pid}_配置別_5セット_問題文_アプリ用.json`",
        f"- Copilot 用: `task/Word_類題_Copilot_project{pid}_配置別.md`",
        "",
        "---",
        "",
        "## 概要",
        "",
        "### 操作種別（教材準拠・変更不可）",
        "",
        "| taskId | 操作 |",
        "|--------|------|",
    ]
    for tid, op in PROJECT_OVERVIEW_OPS.get(pid, []):
        lines.append(f"| {tid} | {op} |")
    lines.append("")

    if pid == 6:
        lines += [
            "### セット間で変えるパラメータ",
            "",
            "| セット | 6-1 | 6-2 | 6-3 | 6-5 | 6-7 |",
            "|--------|-----|-----|-----|-----|-----|",
        ]
        for s in data["sets"]:
            p1, p2, p3, p5, p7 = (
                s["taskParams"]["task6_1"],
                s["taskParams"]["task6_2"],
                s["taskParams"]["task6_3"],
                s["taskParams"]["task6_5"],
                s["taskParams"]["task6_7"],
            )
            op5 = (
                f"列幅{','.join(p5['columnWidthsCm'])}cm+行高{p5['rowHeightCm']}cm"
                if p5.get("operation") == "setColumnWidthsAndRowHeight"
                else "列幅均等"
            )
            lines.append(
                f"| {s['setNo']} {s['theme']} | {p1['rowCount']}×{p1['colCount']} | "
                f"{p2['splitAtRow']}行目 | {p3['splitRow']}行{p3['splitCol']}列→{p3['splitInto']} | {op5} | "
                f"p{p7['page']} {p7['rows']}×{p7['cols']} |"
            )
        lines.append("")

    lines += [
        f"**禁止文字列（教材流用不可）:** {' / '.join(forbidden)}",
        "",
        "### Copilot で docx を作るときの注意（空データ・骨組み対策）",
        "",
        "Copilot にプロンプトだけ渡すと、**タスク用データだけの薄い docx** や**プレースホルダ入り骨組み**になりがちです。",
        "",
        "1. **documentOutline（見出し1〜4＋本文）を先に配置**し、MOS本番相当の文書量にする",
        "2. **貼り付け用データをすべて Word に転記**する（表・タブ区切り・番号段落等を省略しない）",
        "3. **空セル・プレースホルダ・ダミー文字を入れない**",
        "4. **受験者が行う操作は事前に完了させない**",
        "5. Markdown の回答だけで終えず、必ず **.docx ファイルを保存**する",
        f"6. 詳細手順は `task/Word_類題_Copilot_project{pid}_配置別.md` のセッション1プロンプトを使用",
        "",
        "---",
        "",
    ]

    for s in data["sets"]:
        ly = s["layout"]
        cb = s["contentBlocks"]
        headings = " / ".join(ly.get("headings", []))
        lines.append(f"## セット{s['setNo']}：{s['workbook']}")
        lines.append("")
        lines.append(f"**テーマ:** {s['theme']}  ")
        layout_parts = [
            f"余白={ly.get('marginPreset')}",
            f"向き={ly.get('orientation')}",
            f"スタイルセット={ly.get('styleSet')}",
            f"テーマカラー={ly.get('themeColor')}",
        ]
        if ly.get("pageCount"):
            layout_parts.append(f"ページ数={ly['pageCount']}")
        if ly.get("format"):
            layout_parts.append(f"書式={ly['format']}")
        if ly.get("trackChangesEnabled"):
            layout_parts.append("変更履歴=ON")
        if ly.get("compatibilityMode"):
            layout_parts.append("互換モード=有効")
        lines.append(f"**レイアウト:** {' / '.join(layout_parts)}  ")
        lines.append(f"**見出し構成:** {headings}  ")
        if ly.get("tablePlacement"):
            lines.append(f"**表配置:** {ly['tablePlacement']}  ")
        lines.append("")
        lines.append("### 問題文")
        lines.append("")
        lines.append(f"{s['problemStatement']}  ")
        lines.append("問題）  ")
        for t in s["tasks"]:
            lines.append(f"{t}  ")
        lines.append("")
        _append_document_outline_md(lines, cb)
        lines.append("### docx 事前配置（貼り付け用データ）")
        lines.append("")
        lines.append("以下を **すべて** Word 文書に配置してください。骨組み・空欄は不可です。")
        lines.append("")
        _append_paste_blocks_md(lines, pid, s)
        if pid == 6:
            lines.append("### 受験者操作前の状態チェックリスト")
            lines.append("")
            lines.extend(_p6_preop_state(s))
            lines.append("")
        lines.append("---")
        lines.append("")
    return "\n".join(lines)


def write_md_project6(data: dict) -> str:
    return write_md_rich(data)


def write_md(data: dict) -> str:
    pid = data.get("projectId")
    if pid in (6, 7, 8, 9, 10):
        return write_md_rich(data)
    lines = [
        f"# MOSスペシャリスト Word 類題 — Project{pid} 配置別（5セット）",
        "",
        f"元ファイル {data['sourceFile']} の操作をもとに、",
        "レイアウトを差別化した 5 セットの類題文書用問題文です。",
        "",
        "**関連ファイル:**",
        f"- 正本 JSON: `task/類題Json/MOS_Word類題_project{pid}_配置別_5セット_問題文.json`",
        f"- アプリ用 JSON: `task/類題Json/MOS_Word類題_project{pid}_配置別_5セット_問題文_アプリ用.json`",
        f"- Copilot 用: `task/Word_類題_Copilot_project{pid}_配置別.md`",
        "",
    ]
    for s in data["sets"]:
        ly = s["layout"]
        lines.append(f"## セット{s['setNo']}：{s['workbook']}")
        lines.append("")
        lines.append(f"**テーマ:** {s['theme']}  ")
        lines.append(
            f"**レイアウト:** 余白={ly.get('marginPreset')} / 向き={ly.get('orientation')} / "
            f"スタイルセット={ly.get('styleSet')} / テーマカラー={ly.get('themeColor')}"
        )
        lines.append("")
        lines.append(f"{s['problemStatement']}  ")
        lines.append("問題）  ")
        for t in s["tasks"]:
            lines.append(f"{t}  ")
        lines.append("")
    return "\n".join(lines)


def write_app_json(data: dict) -> dict:
    pid = data["projectId"]
    app_sets = []
    for s in data["sets"]:
        items = []
        for i, t in enumerate(s["tasks"], 1):
            desc = re.sub(rf"^タスク{pid}-\d+　", "", t)
            if i == 1:
                desc = s["problemStatement"] + "\n" + desc
            items.append({"taskId": i, "description": desc})
        app_sets.append(
            {
                "setNo": s["setNo"],
                "theme": s["theme"],
                "workbook": s["workbook"],
                "projectId": pid,
                "tasks": items,
            }
        )
    return {
        "sourceFile": data["sourceFile"],
        "variantType": data.get("variantType", "layout"),
        "sets": app_sets,
    }


def _check_unique(values: list, label: str, issues: list[str]) -> None:
    if len(values) != len(set(values)):
        issues.append(f"{label} reused across sets: {values}")


def _collect_param(sets: list, key: str) -> list:
    values = []
    for s in sets:
        for params in s.get("taskParams", {}).values():
            val = params.get(key)
            if val:
                values.append(val)
    return values


def validate_common(data: dict, cfg: dict) -> list[str]:
    issues: list[str] = []
    pid = data["projectId"]
    forbidden = cfg["forbidden"]
    task_count = cfg["task_count"]
    md = write_md(data)

    layout_signatures = []
    formats = []
    theme_colors = []

    for s in data["sets"]:
        if len(s["tasks"]) != task_count:
            issues.append(f"set{s['setNo']}: task count {len(s['tasks'])} != {task_count}")
        if s["workbook"] not in md:
            issues.append(f"set{s['setNo']}: workbook missing in md")
        for t in s["tasks"]:
            if t not in md:
                issues.append(f"set{s['setNo']}: task missing in md: {t[:50]}")

        blob = json.dumps(s, ensure_ascii=False)
        for text in forbidden:
            if text in blob:
                issues.append(f"set{s['setNo']}: forbidden text found: {text}")

        layout = s["layout"]
        layout_signatures.append(
            (layout.get("marginPreset"), layout.get("orientation"), layout.get("styleSet"))
        )
        formats.append(layout.get("format"))
        theme_colors.append(layout.get("themeColor"))

        params = s.get("taskParams", {})
        for key in [f"task{pid}_{i}" for i in range(1, task_count + 1)]:
            if key not in params:
                issues.append(f"set{s['setNo']}: missing {key}")

    if len(set(layout_signatures)) != 5:
        issues.append(f"layout signatures not unique: {layout_signatures}")
    if len(set(formats)) != 5:
        issues.append(f"format not unique: {formats}")
    if len(set(theme_colors)) != 5:
        issues.append(f"themeColor not unique: {theme_colors}")

    return issues


def validate_project_1(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    all_headings, all_start, all_images, all_targets = [], [], [], []
    all_p14_labels, all_p13_styles, all_p14_styles, all_bullet_counts = [], [], [], []

    for s in sets:
        p2 = s["taskParams"].get("task1_2", {})
        p3 = s["taskParams"].get("task1_3", {})
        p4 = s["taskParams"].get("task1_4", {})
        p5 = s["taskParams"].get("task1_5", {})
        if p2.get("heading"):
            all_headings.append(p2["heading"])
        if p2.get("startParagraph"):
            all_start.append(p2["startParagraph"])
        if p3.get("imageName"):
            all_images.append(p3["imageName"])
        if p3.get("targetParagraph"):
            all_targets.append(p3["targetParagraph"])
        if p2.get("paragraphCount") is not None:
            all_bullet_counts.append(p2["paragraphCount"])
        if p3.get("styleUi"):
            all_p13_styles.append(p3["styleUi"])
        label = p4.get("targetLabel") or p4.get("subheading")
        if label:
            all_p14_labels.append(label)
        if p4.get("targetStyle"):
            all_p14_styles.append(p4["targetStyle"])
        if p5.get("targetParagraph"):
            all_targets.append(p5["targetParagraph"])

    image_placements = [s["layout"].get("imagePlacement") for s in sets]
    if len(set(image_placements)) != 5:
        issues.append(f"imagePlacement not unique: {image_placements}")

    for label, values in [
        ("heading", all_headings),
        ("startParagraph", all_start),
        ("imageName", all_images),
        ("targetParagraph", all_targets),
        ("task1_4 targetLabel", all_p14_labels),
        ("task1_3 styleUi", all_p13_styles),
        ("task1_4 targetStyle", all_p14_styles),
    ]:
        _check_unique(values, label, issues)
    return issues


def validate_project_2(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    textbook_color = "青、アクセント1、黒+基本色25％"
    italic_count = 0
    for s in sets:
        p1 = s["taskParams"]["task2_1"]
        p2 = s["taskParams"]["task2_2"]
        p3 = s["taskParams"]["task2_3"]
        p4 = s["taskParams"]["task2_4"]
        p5 = s["taskParams"]["task2_5"]
        if "こちら↓" in p1["cutText"] or "コチラ↓" in p1["cutText"]:
            issues.append(f"set{s['setNo']}: cutText must not use こちら↓ pattern")
        if p2.get("listLevel") == 3:
            issues.append(f"set{s['setNo']}: listLevel must not be 3")
        if p3.get("fontColorUi") == textbook_color:
            issues.append(f"set{s['setNo']}: fontColorUi must not be textbook color")
        if p4.get("shapeLocation") == "documentBottom" or p4.get("shapeType") == "テキストボックス":
            issues.append(f"set{s['setNo']}: shape must not be bottom textbox")
        if p5.get("fontSizePt") == 15:
            issues.append(f"set{s['setNo']}: fontSizePt must not be 15")
        if p5.get("fontStyle") == "斜体":
            italic_count += 1
    if italic_count != 1:
        issues.append(f"task2_5 italic count must be 1, got {italic_count}")

    for label, values in [
        ("cutText", [s["taskParams"]["task2_1"]["cutText"] for s in sets]),
        ("targetHeading", [s["taskParams"]["task2_1"]["targetHeading"] for s in sets]),
        ("task2_2 heading", [s["taskParams"]["task2_2"]["heading"] for s in sets]),
        ("targetText", [s["taskParams"]["task2_3"]["targetText"] for s in sets]),
        ("fontColorUi", [s["taskParams"]["task2_3"]["fontColorUi"] for s in sets]),
        ("shapeType", [s["taskParams"]["task2_4"]["shapeType"] for s in sets]),
        ("fontSizePt", [s["taskParams"]["task2_5"]["fontSizePt"] for s in sets]),
        ("inputText", [s["taskParams"]["task2_4"]["inputText"] for s in sets]),
    ]:
        _check_unique(values, label, issues)
    return issues


def validate_project_3(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    section_headings = [s["taskParams"]["task3_2"]["sectionHeading"] for s in sets]
    column_breaks = [s["taskParams"]["task3_5"]["columnBreakBefore"] for s in sets]
    _check_unique(section_headings, "sectionHeading", issues)
    _check_unique(column_breaks, "columnBreakBefore", issues)

    orientations: list[str] = []
    for s in sets:
        p1 = s["taskParams"]["task3_1"]
        p2 = s["taskParams"]["task3_2"]
        p3 = s["taskParams"]["task3_3"]
        p4 = s["taskParams"]["task3_4"]
        p6 = s["taskParams"]["task3_6"]
        bib = p2["sectionHeading"]

        if p1.get("marginPreset") == "やや狭い":
            issues.append(f"set{s['setNo']}: task3_1 marginPreset must not be やや狭い")
        if "やや狭い" in s["tasks"][0]:
            issues.append(f"set{s['setNo']}: task3-1 text must not use やや狭い")

        if p2.get("breakType") == "nextPage":
            issues.append(f"set{s['setNo']}: task3_2 breakType must not be nextPage")
        if "次のページから" in s["tasks"][1]:
            issues.append(f"set{s['setNo']}: task3-2 text must not use 次のページから")

        if p3.get("sectionHeading") == bib:
            issues.append(
                f"set{s['setNo']}: task3_3 must target body section before bibliography, not {bib}"
            )
        orientations.append(p3.get("orientation", ""))

        cols = p4.get("columnPreset") or str(p4.get("columns", ""))
        if cols == "2" or (p4.get("columns") == 2 and cols in ("2", "")):
            issues.append(f"set{s['setNo']}: task3_4 must not use standard 2-column layout")

        if p6.get("lineSpacing") == 1.3:
            issues.append(f"set{s['setNo']}: task3_6 lineSpacing must not be 1.3")
        if '"1.3"' in s["tasks"][5] or "1.3" in s["tasks"][5].split("行間")[-1][:10]:
            issues.append(f"set{s['setNo']}: task3-6 text must not use 1.3")

    if "縦向き" not in orientations or "横向き" not in orientations:
        issues.append("task3_3: sets must include both 縦向き and 横向き")
    return issues


def _p4_anchor_type(p1: dict) -> str:
    return p1.get("anchorType", "heading")


def validate_project_4(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    anchor_types = [_p4_anchor_type(s["taskParams"]["task4_1"]) for s in sets]
    c1 = [s["taskParams"]["task4_1"]["commentText"] for s in sets]
    h2 = [s["taskParams"]["task4_2"]["heading"] for s in sets]
    r2 = [s["taskParams"]["task4_2"]["replyText"] for s in sets]
    wm = [s["taskParams"]["task4_5"]["watermarkText"] for s in sets]
    ss = [s["taskParams"]["task4_4"]["styleSet"] for s in sets]
    bw = [s["taskParams"]["task4_6"]["borderWidth"] for s in sets]
    for label, values in [
        ("task4_1 anchorType", anchor_types),
        ("task4_1 commentText", c1),
        ("task4_2 heading", h2),
        ("task4_2 replyText", r2),
        ("watermarkText", wm),
        ("styleSet", ss),
        ("borderWidth", bw),
    ]:
        _check_unique(values, label, issues)

    ops = {s["taskParams"]["task4_3"]["operation"] for s in sets}
    if "resolveOnly" not in ops:
        issues.append("task4_3: must include resolveOnly")
    if "deleteOnly" not in ops:
        issues.append("task4_3: must include deleteOnly")
    if "resolveAndDeleteMultiple" not in ops:
        issues.append("task4_3: must include resolveAndDeleteMultiple")

    for s in sets:
        p1 = s["taskParams"]["task4_1"]
        p2 = s["taskParams"]["task4_2"]
        p4 = s["taskParams"]["task4_4"]
        p7 = s["taskParams"]["task4_7"]
        if "確認" in p1.get("commentText", ""):
            issues.append(f"set{s['setNo']}: task4_1 commentText must not contain 確認")
        if "最終確認" in p2.get("replyText", ""):
            issues.append(f"set{s['setNo']}: task4_2 replyText must not contain 最終確認")
        if p4.get("styleSet") == "線（シンプル）":
            issues.append(f"set{s['setNo']}: task4_4 styleSet must not be 線（シンプル）")
        removed = set(p7.get("inspectorRemove", []))
        if removed == {"透かし文字", "ヘッダー・フッター"}:
            issues.append(
                f"set{s['setNo']}: task4_7 must not use textbook inspector items only"
            )
        if "透かし文字とヘッダー・フッター" in s["tasks"][6]:
            issues.append(f"set{s['setNo']}: task4-7 text must not use 透かし文字とヘッダー・フッター")
    return issues


def validate_project_5(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    paragraphs = [s["taskParams"]["task5_1"]["targetParagraph"] for s in sets]
    p_images = [s["taskParams"]["task5_1"]["imageName"] for s in sets]
    anchor_types = [s["taskParams"]["task5_1"].get("anchorType", "paragraphStart") for s in sets]
    wrap1 = [s["taskParams"]["task5_1"].get("wrapType", "") for s in sets]
    wrap2 = [s["taskParams"]["task5_2"].get("wrapType", "") for s in sets]
    arts = [s["taskParams"]["task5_3"].get("artEffect", "") for s in sets]
    blurs = [s["taskParams"]["task5_4"].get("softEdgePt", 0) for s in sets]
    t_images = [s["taskParams"]["task5_5"]["imageName"] for s in sets]
    alts = []
    for s in sets:
        p6 = s["taskParams"]["task5_6"]
        alts.append(p6.get("altText", ""))
    end_images = [s["taskParams"]["task5_8"]["imageName"] for s in sets]
    for label, values in [
        ("targetParagraph", paragraphs),
        ("task5_1 imageName", p_images),
        ("task5_1 anchorType", anchor_types),
        ("task5_1 wrapType", wrap1),
        ("task5_2 wrapType", wrap2),
        ("task5_3 artEffect", arts),
        ("task5_4 softEdgePt", blurs),
        ("task5_5 imageName", t_images),
        ("altText", alts),
        ("endImageName", end_images),
    ]:
        _check_unique(values, label, issues)

    ops8 = {s["taskParams"]["task5_8"].get("operation", "removeBackgroundOnly") for s in sets}
    if "addForegroundMark" not in ops8:
        issues.append("task5_8: must include addForegroundMark")
    wordings7 = [s["taskParams"]["task5_7"].get("wording", "装飾化") for s in sets]
    if wordings7.count("screenReaderHidden") < 3:
        issues.append("task5_7: at least 3 sets must use screenReaderHidden wording")

    no_alt_word = sum(
        1 for s in sets if not s["taskParams"]["task5_6"].get("useAltTextWord", True)
    )
    if no_alt_word < 3:
        issues.append("task5_6: at least 3 sets must not use 代替テキスト in problem text")

    for s in sets:
        p1 = s["taskParams"]["task5_1"]
        p2 = s["taskParams"]["task5_2"]
        p3 = s["taskParams"]["task5_3"]
        p4 = s["taskParams"]["task5_4"]
        p5 = s["taskParams"]["task5_5"]
        if p1.get("wrapType") == "上下":
            issues.append(f"set{s['setNo']}: task5_1 wrapType must not be 上下")
        if p2.get("wrapType") == "狭く":
            issues.append(f"set{s['setNo']}: task5_2 wrapType must not be 狭く")
        if p3.get("artEffect") == "水彩：スポンジ":
            issues.append(f"set{s['setNo']}: task5_3 must not use 水彩：スポンジ")
        if p4.get("softEdgePt") == 25:
            issues.append(f"set{s['setNo']}: task5_4 must not use 25pt blur")
        if p5.get("bevelEffect") == "面取り ハードエッジ":
            issues.append(f"set{s['setNo']}: task5_5 must not use 面取り ハードエッジ")
        if "水彩：スポンジ" in s["tasks"][2]:
            issues.append(f"set{s['setNo']}: task5-3 text must not mention 水彩：スポンジ")
        if "ぼかし25ポイント" in s["tasks"][3]:
            issues.append(f"set{s['setNo']}: task5-4 text must not mention ぼかし25ポイント")
    return issues


def _check_placeholders(blob: str, patterns: list[str], set_no: int, issues: list[str]) -> None:
    for pat in patterns:
        if pat in blob:
            issues.append(f"set{set_no}: placeholder pattern found: {pat}")


def validate_document_outline(data: dict) -> list[str]:
    issues: list[str] = []
    for s in data["sets"]:
        cb = s.get("contentBlocks", {})
        outline = cb.get("documentOutline")
        if not outline:
            issues.append(f"set{s['setNo']}: missing documentOutline")
            continue
        sections = outline.get("sections", [])
        if len(sections) < 3:
            issues.append(f"set{s['setNo']}: documentOutline needs at least 3 level-1 sections")
        max_lv = _max_outline_level(sections)
        min_lv = outline.get("minHeadingLevel", 4)
        if max_lv < min_lv:
            issues.append(f"set{s['setNo']}: documentOutline max level {max_lv} < {min_lv}")
        if not outline.get("title"):
            issues.append(f"set{s['setNo']}: documentOutline missing title")
    return issues


def validate_content_blocks(data: dict, cfg: dict) -> list[str]:
    issues: list[str] = []
    required = cfg.get("content_required", [])
    if not required:
        return issues
    patterns = data.get("variantRules", {}).get("forbiddenPlaceholderPatterns", [])
    for s in data["sets"]:
        cb = s.get("contentBlocks")
        if not cb:
            issues.append(f"set{s['setNo']}: missing contentBlocks")
            continue
        for key in required:
            if key not in cb:
                issues.append(f"set{s['setNo']}: contentBlocks missing {key}")
        blob = json.dumps(cb, ensure_ascii=False)
        _check_placeholders(blob, patterns, s["setNo"], issues)
    if "documentOutline" in required:
        issues.extend(validate_document_outline(data))
    return issues


def validate_project_6(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    shapes, split_rows, split_specs, specs_67, ops_65 = [], [], [], [], []
    for s in sets:
        p1 = s["taskParams"]["task6_1"]
        p2 = s["taskParams"]["task6_2"]
        p3 = s["taskParams"]["task6_3"]
        p4 = s["taskParams"]["task6_4"]
        p5 = s["taskParams"]["task6_5"]
        p7 = s["taskParams"]["task6_7"]
        cb = s.get("contentBlocks", {})
        tab_rows = cb.get("task6_1_tabRows", [])
        row_count = p1.get("rowCount", 0)
        col_count = p1.get("colCount", 0)
        if len(tab_rows) != row_count:
            issues.append(
                f"set{s['setNo']}: task6_1_tabRows rows {len(tab_rows)} != rowCount {row_count}"
            )
        for i, row in enumerate(tab_rows):
            if len(row) != col_count:
                issues.append(
                    f"set{s['setNo']}: task6_1_tabRows row{i+1} cols {len(row)} != colCount {col_count}"
                )
        if p2.get("splitAtRow") == 5:
            issues.append(f"set{s['setNo']}: splitAtRow must not be 5 (textbook value)")
        if row_count == 7 and col_count == 2:
            issues.append(f"set{s['setNo']}: task6_1 must not be 7x2 (textbook shape)")
        if p3.get("splitRow") == 1 and p3.get("splitCol") == 2 and p3.get("splitInto") == 3:
            issues.append(f"set{s['setNo']}: task6_3 must not be row1col2 splitInto 3 (textbook)")
        if p7.get("page") == 2 and p7.get("rows") == 13 and p7.get("cols") == 3:
            issues.append(f"set{s['setNo']}: task6_7 must not be page2 13x3 (textbook)")
        labels = p7.get("headerCells", [])
        if labels == ["最高収益", "平均収益", "合計収益"]:
            issues.append(f"set{s['setNo']}: task6_7 labels must not be 最高収益/平均収益/合計収益")
        if "オレンジ" in p4.get("textEffectFill", ""):
            issues.append(f"set{s['setNo']}: task6_4 must not use orange accent2")
        tbl = cb.get("task6_4_table", {})
        if len(tbl.get("rows", [])) < 4:
            issues.append(f"set{s['setNo']}: task6_4_table needs at least 4 data rows")
        if p5.get("operation") == "setColumnWidthsAndRowHeight":
            widths = p5.get("columnWidthsCm", [])
            if len(widths) != len(tbl.get("headers", [])):
                issues.append(f"set{s['setNo']}: columnWidthsCm count must match table columns")
            if not p5.get("rowHeightCm"):
                issues.append(f"set{s['setNo']}: rowHeightCm required for setColumnWidthsAndRowHeight")
        trailing = cb.get("task6_7_trailingBlankPage")
        if trailing != p7.get("page"):
            issues.append(
                f"set{s['setNo']}: task6_7_trailingBlankPage {trailing} != task6_7.page {p7.get('page')}"
            )
        shapes.append((row_count, col_count))
        split_rows.append(p2.get("splitAtRow"))
        split_specs.append((p3.get("splitRow"), p3.get("splitCol"), p3.get("splitInto")))
        specs_67.append((p7.get("page"), p7.get("rows"), p7.get("cols"), tuple(labels)))
        ops_65.append(p5.get("operation"))
    for label, values in [
        ("table6_1_shape", shapes),
        ("splitAtRow", split_rows),
        ("splitCellSpec", split_specs),
        ("heading6_1", [s["taskParams"]["task6_1"]["heading"] for s in sets]),
        ("startText", [s["taskParams"]["task6_1"]["startText"] for s in sets]),
        ("heading6_4", [s["taskParams"]["task6_4"]["heading"] for s in sets]),
        ("textEffectFill", [s["taskParams"]["task6_4"]["textEffectFill"] for s in sets]),
        ("task6_7_spec", specs_67),
        ("headerCells6_7", [tuple(s["taskParams"]["task6_7"]["headerCells"]) for s in sets]),
    ]:
        _check_unique(values, label, issues)
    if ops_65.count("setColumnWidthsAndRowHeight") < 2:
        issues.append("task6_5: at least 2 sets must use setColumnWidthsAndRowHeight")
    if ops_65.count("distributeColumns") < 2:
        issues.append("task6_5: at least 2 sets must use distributeColumns")
    return issues


def validate_project_7(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    for label, values in [
        ("propertyValue", [s["taskParams"]["task7_2"]["propertyValue"] for s in sets]),
        ("propertyName", [s["taskParams"]["task7_2"]["propertyName"] for s in sets]),
        ("headerText", [s["taskParams"]["task7_3"]["headerText"] for s in sets]),
        ("firstPageHeaderText", [s["taskParams"]["task7_4"]["headerText"] for s in sets]),
        ("footerText", [s["taskParams"]["task7_6"]["footerText"] for s in sets]),
        ("txtBaseName", [s["taskParams"]["task7_7"]["txtBaseName"] for s in sets]),
        ("readPassword", [s["taskParams"]["task7_8"]["readPassword"] for s in sets]),
    ]:
        _check_unique(values, label, issues)
    for s in sets:
        p2 = s["taskParams"]["task7_2"]
        if p2.get("propertyName") == "会社名" or p2.get("propertyKey") == "Company":
            issues.append(f"set{s['setNo']}: task7_2 must not use 会社名/Company")
        if p2.get("propertyValue") in ("ラビット出版", "インテグラル"):
            issues.append(f"set{s['setNo']}: task7_2 propertyValue is textbook value")
        if s["taskParams"]["task7_8"]["readPassword"] == "abc":
            issues.append(f"set{s['setNo']}: readPassword must not be abc")
        if s["taskParams"]["task7_5"].get("footerType") != "pageNumber":
            issues.append(f"set{s['setNo']}: task7_5 must be pageNumber footer")
        body = s.get("contentBlocks", {}).get("bodyParagraphs", [])
        if len(body) < 3:
            issues.append(f"set{s['setNo']}: bodyParagraphs needs at least 3 paragraphs")
    if len(set(s["taskParams"]["task7_2"]["propertyName"] for s in sets)) != 5:
        issues.append("task7_2: propertyName must vary across all 5 sets")
    return issues


def validate_project_8(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    textbook_smart_art = ("基本ベン図",)
    for s in sets:
        p1 = s["taskParams"]["task8_1"]
        p3 = s["taskParams"]["task8_3"]
        p4 = s["taskParams"]["task8_4"]
        p5 = s["taskParams"]["task8_5"]
        p6 = s["taskParams"]["task8_6"]
        p7 = s["taskParams"]["task8_7"]
        h0 = s["layout"]["headings"][0]
        if p1.get("insertInShape"):
            issues.append(f"set{s['setNo']}: task8_1 must not insert TOC in shape")
        if p1.get("tocPlacement") == "belowSubtitleInShape":
            issues.append(f"set{s['setNo']}: task8_1 must not use belowSubtitleInShape")
        if "って何" in h0:
            issues.append(f"set{s['setNo']}: first heading must not use って何 pattern")
        if p3.get("footnoteStartNumber") == "①":
            issues.append(f"set{s['setNo']}: footnoteStartNumber must not be ①")
        if p4.get("specialChar") == "§":
            issues.append(f"set{s['setNo']}: specialChar must not be §")
        if p5.get("smartArtType") in textbook_smart_art:
            issues.append(f"set{s['setNo']}: smartArtType must not be 基本ベン図")
        if p6.get("smartArtColor") == "カラフル":
            issues.append(f"set{s['setNo']}: smartArtColor must not be カラフル")
        if p7.get("linkShapePlacement") == "documentBottom":
            issues.append(f"set{s['setNo']}: link must not be at document bottom")
        if p7.get("linkShapeText") == "TOP":
            issues.append(f"set{s['setNo']}: linkShapeText must not be TOP")
        if p7.get("linkType") == "documentTop" and p7.get("linkShapePlacement") == "documentBottom":
            issues.append(f"set{s['setNo']}: documentTop link must not use documentBottom placement")
        cb = s.get("contentBlocks", {})
        if not cb.get("linkShape"):
            issues.append(f"set{s['setNo']}: missing linkShape in contentBlocks")
        if cb.get("bottomShapeText"):
            issues.append(f"set{s['setNo']}: bottomShapeText must not be used (use linkShape)")
        bib = cb.get("bibliographyEntries", [])
        if len(bib) < 3:
            issues.append(f"set{s['setNo']}: bibliographyEntries needs at least 3 entries")
    for label, values in [
        ("subtitle", [s["taskParams"]["task8_1"]["subtitle"] for s in sets]),
        ("tocPlacement", [s["taskParams"]["task8_1"]["tocPlacement"] for s in sets]),
        ("footnoteAnchor", [s["taskParams"]["task8_2"]["footnoteAnchor"] for s in sets]),
        ("footnoteText", [s["taskParams"]["task8_2"]["footnoteText"] for s in sets]),
        ("footnoteStartNumber", [s["taskParams"]["task8_3"]["footnoteStartNumber"] for s in sets]),
        ("bibliographyHeading", [s["taskParams"]["task8_4"]["heading"] for s in sets]),
        ("specialChar", [s["taskParams"]["task8_4"]["specialChar"] for s in sets]),
        ("smartArtType", [s["taskParams"]["task8_5"]["smartArtType"] for s in sets]),
        ("smartArtLabels", [tuple(s["taskParams"]["task8_5"]["smartArtLabels"]) for s in sets]),
        ("smartArtColor", [s["taskParams"]["task8_6"]["smartArtColor"] for s in sets]),
        ("linkShapeText", [s["taskParams"]["task8_7"]["linkShapeText"] for s in sets]),
        ("linkTarget", [
            s["taskParams"]["task8_7"].get("linkUrl")
            or s["taskParams"]["task8_7"].get("linkTargetHeading")
            or s["taskParams"]["task8_7"]["linkType"]
            for s in sets
        ]),
    ]:
        _check_unique(values, label, issues)
    link_types = [s["taskParams"]["task8_7"]["linkType"] for s in sets]
    if "url" not in link_types:
        issues.append("task8_7: at least one set must use url linkType")
    if "heading" not in link_types and "documentEnd" not in link_types:
        issues.append("task8_7: should include heading or documentEnd linkType")
    return issues


def validate_project_9(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    comma_count = 0
    tc_ops: list[str] = []
    for s in sets:
        if s["taskParams"]["task9_6"]["trackChangesPassword"] == "654":
            issues.append(f"set{s['setNo']}: trackChangesPassword must not be 654")
        samples = s.get("contentBlocks", {}).get("trackChangesSamples", [])
        if len(samples) < 2:
            issues.append(f"set{s['setNo']}: trackChangesSamples needs at least 2 samples")
        tbl = s.get("contentBlocks", {}).get("task9_3_table", {})
        if len(tbl.get("rows", [])) < 4:
            issues.append(f"set{s['setNo']}: task9_3_table needs at least 4 data rows")
        p1 = s["taskParams"]["task9_1"]
        if p1.get("restartAt", 1) == 1:
            issues.append(f"set{s['setNo']}: task9_1 restartAt must not be 1")
        p2 = s["taskParams"]["task9_2"]
        if p2.get("separator") == "comma":
            comma_count += 1
        elif p2.get("operation") == "convertTableToCommaText":
            comma_count += 1
        p3 = s["taskParams"]["task9_3"]
        if p3.get("marginSide", "right") == "right":
            issues.append(f"set{s['setNo']}: task9_3 marginSide must not be right")
        if p3.get("marginMm", 4) == 4:
            issues.append(f"set{s['setNo']}: task9_3 marginMm must not be 4")
        p4 = s["taskParams"]["task9_4"]
        total_cols = {"売上合計", "総計", "出荷計", "売上計", "受講計"}
        if p4.get("sortColumn") in total_cols and p4.get("sortOrder") == "desc":
            issues.append(
                f"set{s['setNo']}: task9_4 must not use total column descending sort"
            )
        p5 = s["taskParams"]["task9_5"]
        op = p5.get("operation", "acceptAllAndStopTracking")
        if op == "acceptAllAndStopTracking":
            issues.append(f"set{s['setNo']}: task9_5 must not use acceptAllAndStopTracking")
        tc_ops.append(op)
    if comma_count != 2:
        issues.append(f"task9_2: comma separator must appear in exactly 2 sets, got {comma_count}")
    required_tc = {"acceptAll", "revertAll", "rejectAllAndStop"}
    if not required_tc.issubset(set(tc_ops)):
        issues.append(f"task9_5: must include all operations {required_tc}, got {tc_ops}")
    for label, values in [
        ("numberedParagraph", [s["taskParams"]["task9_1"]["numberedParagraph"] for s in sets]),
        ("table9_2Heading", [s["taskParams"]["task9_2"]["heading"] for s in sets]),
        ("table9_3Heading", [s["taskParams"]["task9_3"]["tableHeading"] for s in sets]),
        ("sortColumn", [s["taskParams"]["task9_4"]["sortColumn"] for s in sets]),
        ("trackChangesPassword", [s["taskParams"]["task9_6"]["trackChangesPassword"] for s in sets]),
        ("restartAt", [s["taskParams"]["task9_1"]["restartAt"] for s in sets]),
        ("marginMm", [s["taskParams"]["task9_3"]["marginMm"] for s in sets]),
    ]:
        _check_unique(values, label, issues)
    return issues


def validate_project_10(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    format_types: list[str] = []
    for s in sets:
        p1 = s["taskParams"]["task10_1"]
        search = p1["searchText"]
        replace = p1["replaceText"]
        occ = p1.get("occurrenceCount", 0)
        body = s.get("contentBlocks", {}).get("bodyWithSearchTerms", [])
        last = s.get("contentBlocks", {}).get("lastTwoLines", [])
        all_body = body + last
        actual = sum(p.count(search) for p in all_body)
        if occ < 2:
            issues.append(f"set{s['setNo']}: occurrenceCount must be >= 2")
        if actual < occ:
            issues.append(
                f"set{s['setNo']}: body text has {actual} occurrences of "
                f"{search}, expected {occ}"
            )
        if search == replace:
            issues.append(f"set{s['setNo']}: searchText and replaceText must differ")
        if p1.get("searchText") == "ウイルス":
            issues.append(f"set{s['setNo']}: searchText must not be ウイルス")
        p3 = s["taskParams"]["task10_3"]
        if p3.get("bulletFont", "").lower() == "webdings":
            issues.append(f"set{s['setNo']}: task10_3 bulletFont must not be Webdings")
        if str(p3.get("bulletCharCode", "")) == "120":
            issues.append(f"set{s['setNo']}: task10_3 bulletCharCode must not be 120")
        p4 = s["taskParams"]["task10_4"]
        ft = p4.get("formatType", p4.get("italicTarget") and "italic")
        if ft == "italic" or "italicTarget" in p4:
            issues.append(f"set{s['setNo']}: task10_4 must not use italic")
        if p4.get("formatTarget") == "事例":
            issues.append(f"set{s['setNo']}: formatTarget must not be 事例")
        target = p4.get("formatTarget", "")
        fmt_count = sum(p.count(target) for p in all_body)
        if target and fmt_count < 2:
            issues.append(
                f"set{s['setNo']}: formatTarget {target} needs >= 2 occurrences in body"
            )
        format_types.append(ft or "")
        p5 = s["taskParams"]["task10_5"]
        if p5.get("paperSize") == "B5":
            issues.append(f"set{s['setNo']}: paperSize must not be B5")
        lst = p5.get("lineSpacingType", "")
        if lst == "1.6行":
            issues.append(f"set{s['setNo']}: lineSpacingType must not be 1.6行")
        if p5.get("lineSpacing") == 1.6 and not lst:
            issues.append(f"set{s['setNo']}: line spacing must not be 1.6")
        if s["taskParams"]["task10_2"]["bulletImageName"] == "PCアイコン":
            issues.append(f"set{s['setNo']}: bulletImageName must not be PCアイコン")
    required_formats = {"bold", "color", "fontSize"}
    if not required_formats.issubset(set(format_types)):
        issues.append(f"task10_4: must include format types {required_formats}, got {format_types}")
    for label, values in [
        ("searchText", [s["taskParams"]["task10_1"]["searchText"] for s in sets]),
        ("replaceText", [s["taskParams"]["task10_1"]["replaceText"] for s in sets]),
        ("heading10_2", [s["taskParams"]["task10_2"]["heading"] for s in sets]),
        ("bulletImageName", [s["taskParams"]["task10_2"]["bulletImageName"] for s in sets]),
        ("heading10_3", [s["taskParams"]["task10_3"]["heading"] for s in sets]),
        ("formatTarget", [s["taskParams"]["task10_4"]["formatTarget"] for s in sets]),
        ("bulletCharCode", [s["taskParams"]["task10_3"]["bulletCharCode"] for s in sets]),
        ("paperSize", [s["taskParams"]["task10_5"]["paperSize"] for s in sets]),
    ]:
        _check_unique(values, label, issues)
    return issues


PROJECT_VALIDATORS = {
    1: validate_project_1,
    2: validate_project_2,
    3: validate_project_3,
    4: validate_project_4,
    5: validate_project_5,
    6: validate_project_6,
    7: validate_project_7,
    8: validate_project_8,
    9: validate_project_9,
    10: validate_project_10,
}


def validate(data: dict, cfg: dict) -> list[str]:
    issues = validate_common(data, cfg)
    issues.extend(validate_content_blocks(data, cfg))
    pid = data["projectId"]
    validator = PROJECT_VALIDATORS.get(pid)
    if validator:
        issues.extend(validator(data))
    return issues


def build_project(project_id: int) -> list[str]:
    path = json_path(project_id)
    if not path.exists():
        return [f"JSON not found: {path}"]
    data = json.loads(path.read_text(encoding="utf-8"))
    cfg = PROJECT_CONFIG[project_id]
    issues = validate(data, cfg)
    if issues:
        return issues
    md_path(project_id).write_text(write_md(data), encoding="utf-8")
    app_json_path(project_id).write_text(
        json.dumps(write_app_json(data), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return []


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--project", type=int, action="append", help="project id (repeatable)")
    parser.add_argument("--all", action="store_true")
    parser.add_argument("--validate-only", action="store_true")
    args = parser.parse_args()
    ids = list(PROJECT_CONFIG.keys()) if args.all else (args.project or [])
    if not ids:
        parser.print_help()
        sys.exit(1)
    failed = False
    for pid in ids:
        if args.validate_only:
            data = json.loads(json_path(pid).read_text(encoding="utf-8"))
            issues = validate(data, PROJECT_CONFIG[pid])
        else:
            issues = build_project(pid)
        if issues:
            print(f"Word Project{pid}: FAIL", issues)
            failed = True
        else:
            print(f"Word Project{pid}: OK")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
