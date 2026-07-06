"""Generate Copilot MD from Word variant problem JSON."""
import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
TOOLS = Path(__file__).resolve().parent
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
from word_document_outline import render_outline_md


def task_params_table(s: dict, pid: int) -> str:
    rows = []
    for i in range(1, 20):
        key = f"task{pid}_{i}"
        if key not in s.get("taskParams", {}):
            break
        p = s["taskParams"][key]
        parts = [f"{k}={v}" for k, v in p.items()]
        rows.append(f"| {pid}-{i} | {', '.join(parts)} |")
    header = "| taskId | パラメータ |"
    sep = "|--------|------------|"
    return "\n".join([header, sep] + rows)


def layout_table(s: dict) -> str:
    ly = s["layout"]
    headings = " / ".join(ly.get("headings", []))
    lines = [
        "| 項目 | 値 |",
        "|------|-----|",
        f"| セクション数 | {ly.get('sectionCount')} |",
        f"| 見出し構成 | {headings} |",
        f"| 余白 | {ly.get('marginPreset')} |",
        f"| 印刷の向き | {ly.get('orientation')} |",
        f"| スタイルセット | {ly.get('styleSet')} |",
        f"| テーマカラー | {ly.get('themeColor')} |",
        f"| 書式傾向 | {ly.get('format')} |",
        f"| 画像配置 | {ly.get('imagePlacement')} |",
        f"| 改ページ位置 | {ly.get('pageBreakAfter')} |",
        f"| 印刷タイトル行 | {ly.get('printTitleRow')} |",
        f"| 段落数（目安） | {ly.get('paragraphCount')} |",
    ]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Project 1 handler
# ---------------------------------------------------------------------------

def _task13_requirement_p1(p3: dict) -> list[str]:
    anchor = p3.get("anchor", "")
    style = p3.get("styleUi", "")
    if anchor == "leftOfImage":
        return [
            f"   - 画像「{p3['imageName']}」を配置し、その**左**の段落に以下の文を配置",
            f"   - 段落文: 「{p3['targetParagraph']}」",
            f"   - スタイル「{style}」は**未設定**（受験者が設定）",
        ]
    if anchor == "rightOfImage":
        return [
            f"   - 画像「{p3['imageName']}」を配置し、その**右**の段落に以下の文を配置",
            f"   - 段落文: 「{p3['targetParagraph']}」",
            f"   - スタイル「{style}」は**未設定**",
        ]
    if anchor == "belowImage":
        return [
            f"   - 画像「{p3['imageName']}」を配置し、その**下**の段落に以下の文を配置",
            f"   - 段落文: 「{p3['targetParagraph']}」",
            f"   - スタイル「{style}」は**未設定**",
        ]
    if anchor == "underHeading":
        return [
            f"   - 見出し「{p3['heading']}」の直下に段落「{p3['targetParagraph']}」を配置",
            f"   - スタイル「{style}」は**未設定**（画像は不要）",
        ]
    return [f"   - 段落「{p3.get('targetParagraph', '')}」にスタイル「{style}」未設定"]


def _task14_requirement_p1(p4: dict) -> list[str]:
    ttype = p4.get("targetType", "")
    label = p4.get("targetLabel", "")
    style = p4.get("targetStyle", "")
    if ttype == "subheading":
        return [
            f"   - 見出し「{p4['parentHeading']}」内に小見出し「{label}」を配置",
            f"   - 小見出しは「{style}」**以外**のスタイル（受験者が「{style}」に変更）",
        ]
    if ttype == "heading":
        return [
            f"   - 見出し「{label}」を配置（「{style}」以外のスタイル）",
            f"   - 受験者が「{style}」に変更",
        ]
    if ttype == "bodyParagraph":
        parent = p4.get("parentHeading", "")
        prefix = f"見出し「{parent}」内の" if parent else ""
        return [
            f"   - {prefix}本文段落「{label}」を配置（「{style}」以外）",
            f"   - 受験者がスタイルを「{style}」に変更",
        ]
    return [f"   - 「{label}」を「{style}」以外で配置"]


def variation_table_p1(data: dict, pid: int) -> str:
    rows_11, rows_12, rows_13, rows_14 = [], [], [], []
    for s in data["sets"]:
        p1 = s["taskParams"][f"task{pid}_1"]
        p2 = s["taskParams"][f"task{pid}_2"]
        p3 = s["taskParams"][f"task{pid}_3"]
        p4 = s["taskParams"][f"task{pid}_4"]
        step = "／".join(p1.get("steps", []))
        rows_11.append(f"| {s['setNo']} {s['theme']} | {step} | 初期={p1.get('initialState', '—')} |")
        rows_12.append(
            f"| {s['setNo']} {s['theme']} | {p2.get('paragraphCount')}段落 | "
            f"{p2.get('heading')} → {p2.get('startParagraph')} |"
        )
        anchor = p3.get("anchor", "—")
        style = p3.get("styleUi", "—")
        loc = p3.get("imageName") or p3.get("heading", "—")
        rows_13.append(f"| {s['setNo']} {s['theme']} | {anchor} | {style} | {loc} |")
        rows_14.append(
            f"| {s['setNo']} {s['theme']} | {p4.get('targetType', '—')} | "
            f"{p4.get('targetLabel', '—')} | → {p4.get('targetStyle', '—')} |"
        )
    return "\n".join(
        [
            "**タスク1-1 の操作（セットごとに異なる）:**",
            "",
            "| セット | 操作 | 初期状態 |",
            "|--------|------|----------|",
            *rows_11,
            "",
            "**タスク1-2 の箇条書き段落数（セットごとに異なる）:**",
            "",
            "| セット | 段落数 | 見出し → 開始段落 |",
            "|--------|--------|-------------------|",
            *rows_12,
            "",
            "**タスク1-3 のスタイル適用（セットごとに異なる）:**",
            "",
            "| セット | 配置 | スタイル | 対象 |",
            "|--------|------|----------|------|",
            *rows_13,
            "",
            "**タスク1-4 のスタイル変更（セットごとに異なる）:**",
            "",
            "| セット | 対象種別 | 対象文字列 | 変更先 |",
            "|--------|----------|------------|--------|",
            *rows_14,
            "",
        ]
    )


def session1_p1(s: dict, pid: int, forbidden: list[str]) -> str:
    ly = s["layout"]
    p1 = s["taskParams"][f"task{pid}_1"]
    p2 = s["taskParams"][f"task{pid}_2"]
    p3 = s["taskParams"][f"task{pid}_3"]
    p4 = s["taskParams"][f"task{pid}_4"]
    p5 = s["taskParams"][f"task{pid}_5"]
    headings = " / ".join(ly.get("headings", []))
    forbidden_line = "、".join(forbidden[:5]) + " 等"
    task_blocks = [
        "1. **1-1 編集記号**",
        f"   - 初期状態: 編集記号「{p1.get('initialState', '表示')}」",
        f"   - 受験者操作: 「{'／'.join(p1.get('steps', []))}」のみ",
        "",
        "2. **1-2 箇条書き**",
        f"   - 見出し「{p2['heading']}」の直下",
        f"   - 「{p2['startParagraph']}」から連続 **{p2['paragraphCount']}段落**（いずれも未箇条書き）",
        "",
        "3. **1-3 スタイル適用**",
        *_task13_requirement_p1(p3),
        "",
        "4. **1-4 スタイル変更**",
        *_task14_requirement_p1(p4),
        "",
        "5. **1-5 書式クリア**",
        f"   - 見出し「{p5['heading']}」の下に「{p5['targetParagraph']}」を配置",
        f"   - 事前書式: {', '.join(p5.get('prefilledFormat', []))} を付与済み",
    ]
    return _session1_shell(s, pid, forbidden_line, headings, ly, task_blocks)


# ---------------------------------------------------------------------------
# Project 2 handler
# ---------------------------------------------------------------------------

P2_SHAPE_INSERT_HINTS = {
    "吹き出し": "挿入 → 図形 → 吹き出し → 角丸四角形の吹き出し",
    "楕円": "挿入 → 図形 → 基本図形 → 楕円",
    "五角形": "挿入 → 図形 → 基本図形 → 五角形",
    "矢印": "挿入 → 図形 → ブロック矢印 → 右向き矢印",
    "雲形": "挿入 → 図形 → 基本図形 → 雲",
}


def _p2_shape_copilot_block(p4: dict, p5: dict) -> list[str]:
    shape_type = p4.get("shapeType", "図形")
    insert_path = P2_SHAPE_INSERT_HINTS.get(shape_type, f"挿入 → 図形 → {shape_type}")
    anchor = p4.get("anchorHeading", "")
    loc = p4.get("shapeLocation", "underHeading")
    loc_ja = f"見出し「{anchor}」の横" if loc == "besideHeading" else f"見出し「{anchor}」の直下"
    wrap = "四角形" if loc == "besideHeading" else "行内"
    return [
        "5. **2-4 / 2-5 図形（Copilot への作成命令 — 最重要）**",
        "",
        "   次の手順で **Word ネイティブの図形** を挿入すること。画像・イラストで代用しない。",
        "",
        f"   【挿入手順】{insert_path}",
        f"   【配置】{loc_ja}（文字列の折り返し: {wrap}）",
        "   【初期状態】図形内は空、または仮の1文字のみ（2-4の完成文は入れない）",
        f"   【2-4で入力する文（今は入れない）】「{p4.get('inputText', '')}」",
        f"   【2-5で設定する書式（今は設定しない）】{p5.get('fontSizePt')}pt・{p5.get('fontStyle')}",
        "",
        "   【Copilot への命令文（このまま実行）】",
        f"   > Word文書に、{insert_path.split('→')[-1].strip()}を挿入してください。",
        f"   > 配置は見出し「{anchor}」の{'横' if loc == 'besideHeading' else '下'}です。",
        "   > 挿入 → 画像 は使わないでください。ストック画像・生成画像・スクリーンショットも禁止です。",
        "   > 図形をクリックすると「図形の書式」タブが表示され、図形内に文字を入力できること。",
        "   > 図形内の文字は空のまま（または仮の1文字）にしてください。",
        "",
        "   【禁止】",
        "   - 画像（Picture）・アイコン・SmartArt・SVG貼り付けで図形を代用すること",
        "   - 文書最下部の四角いテキストボックスのみを使うこと",
        "   - 2-4の完成テキスト・2-5のフォント書式を事前設定すること",
        "",
        "   【確認方法】図形選択時にリボンが「図形の書式」であること（「図の形式」なら画像なので削除して差し替え）",
        "",
    ]


def usage_section_p2() -> list[str]:
    return [
        "## 0. 使い方（3セッション × 1セット）",
        "",
        "**重要（Project 2・図形）:** Copilot は図形を画像として出力しがちです。",
        "セッション1の **「2-4 / 2-5 図形」ブロック内の Copilot 命令文** を必ず実行し、",
        "「図形の書式」タブで編集できるオブジェクトになっているか確認してください。",
        "",
        "```",
        "1チャット = 1セット（§1〜§5 のいずれか1節のみ使用）",
        "      ↓",
        "セッション1 → 該当 § の「セッション1」プロンプト（本文＋図形を Word 上で作成）",
        "      ↓  画像になった場合は § の「セッション1b」で図形に差し替え",
        "      ↓  docx を保存",
        "セッション2 → 同 § の「セッション2」プロンプト（docx 添付）",
        "      ↓",
        "セッション3 → 同 § の「セッション3」プロンプト（検証済みパラメータ表を貼る）",
        "```",
        "",
        "### 図形が画像になったときの対処",
        "",
        "1. 画像を選択 → Delete",
        "2. セッション1b のプロンプトを Copilot に貼る（または手動で 挿入 → 図形）",
        "3. セッション2 で「図形の書式」タブを確認",
        "",
        "---",
        "",
    ]


def variation_table_p2(data: dict, pid: int) -> str:
    rows = []
    for s in data["sets"]:
        p1 = s["taskParams"][f"task{pid}_1"]
        p2 = s["taskParams"][f"task{pid}_2"]
        p3 = s["taskParams"][f"task{pid}_3"]
        p4 = s["taskParams"][f"task{pid}_4"]
        p5 = s["taskParams"][f"task{pid}_5"]
        rows.append(
            f"| {s['setNo']} {s['theme']} | {p1['cutText']} | L{p2['listLevel']} | "
            f"{p3['fontColorUi']} | {p4['shapeType']} | {p5['fontSizePt']}pt {p5['fontStyle']} |"
        )
    return "\n".join(
        [
            "**タスク2-1〜2-5 のパラメータ（セットごとに異なる）:**",
            "",
            "| セット | 切り取り元 | 2-2レベル | 2-3文字色 | 2-4図形 | 2-5書式 |",
            "|--------|-----------|----------|----------|---------|---------|",
            *rows,
            "",
            "**注意:** 切り取り元は「こちら↓」形式不可。2-2レベルは3以外。",
            "2-3色は「ブルーグレー、テキスト2」+白基本色（80/60/40％）、「濃い赤」、「ゴールド、アクセント4」+白基本色、「薄い青」から選出。斜体は1セットのみ。",
            "",
        ]
    )


def session1_p2(s: dict, pid: int, forbidden: list[str]) -> str:
    ly = s["layout"]
    p1 = s["taskParams"][f"task{pid}_1"]
    p2 = s["taskParams"][f"task{pid}_2"]
    p3 = s["taskParams"][f"task{pid}_3"]
    p4 = s["taskParams"][f"task{pid}_4"]
    p5 = s["taskParams"][f"task{pid}_5"]
    headings = " / ".join(ly.get("headings", []))
    forbidden_line = "、".join(forbidden[:5]) + " 等"
    task_blocks = [
        "1. **2-1 切り取り貼り付け**",
        f"   - 本文段落に「{p1['cutText']}」を配置（**こちら↓形式は使わない**）",
        f"   - 見出し「{p1['targetHeading']}」の下にリンク段落（{p1.get('linkParagraphContains', 'URL')}含む）",
        f"   - 切り取り文字列は見出し下・リンク直上に**未配置**",
        "",
        "2. **2-2 箇条書きレベル**",
        f"   - 見出し「{p2['heading']}」の下に箇条書き **{p2['bulletCount']}項目**（レベル{p2.get('initialLevel', 1)}）",
        f"   - レベル{p2['listLevel']}は**未設定**（受験者が変更。レベル3は使用しない）",
        "",
        "3. **2-3 文字色**",
        f"   - 「{p3['targetText']}」を配置（色「{p3['fontColorUi']}」は**未設定**）",
        "",
        *_p2_shape_copilot_block(p4, p5),
    ]
    return _session1_shell_p2(s, pid, forbidden_line, headings, ly, task_blocks)


def _session1_shell_p2(s, pid, forbidden_line, headings, ly, task_blocks):
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 365 演習アプリ向けの類題設計者です。",
            "**Microsoft Word 上で** docx を直接作成・編集してください（Markdownだけの回答は不可）。",
            f"project{pid} に相当する演習文書を **1冊** 作成します。",
            "問題文テキストは書きません。レイアウトとタスク用データの配置が目的です。",
            "",
            f"【添付】Word問題文一覧_操作手順.csv（projectId={pid}）",
            f"【文書名】{s['workbook']}",
            f"【テーマ】{s['theme']}（{s['problemStatement'].replace('あなたは', '').replace('を作成しています。', '')}）",
            "",
            "【レイアウト仕様 — 必ず遵守】",
            layout_table(s),
            "",
            f"【見出し構成】{headings}",
            f"【段落数目安】{ly.get('paragraphCount')}段落（2〜3ページ相当）",
            "",
            "【作成手順 — この順で Word 上で実行】",
            "",
            "■ ステップA: 見出しと本文",
            "  - 上記見出しを順に配置し、テーマに沿った本文段落を追加",
            "  - スタイルセット・テーマカラーはレイアウト仕様に合わせる",
            "",
            "■ ステップB: タスク用データ（下記要件どおり）",
            "",
            *task_blocks,
            "",
            "【画像に関する厳守事項】",
            "  - 2-4/2-5 の図形は **挿入→図形** のみ。挿入→画像・Copilot画像生成は禁止",
            "  - 装飾用の画像・イラスト・アイコンは文書に入れない",
            "",
            "【禁止】",
            f"  - 教材の禁止文字列（{forbidden_line}）",
            "  - 問題文の生成",
            "  - 受験者操作（切り取り貼り付け・色変更・図形入力・書式）の事前完了",
            "",
            "【出力】",
            f"1. ファイル名「{s['workbook']}」で docx を保存",
            "2. レイアウト仕様書（Markdown）: 見出し一覧・切り取り元段落・図形種別・配置",
            "3. 図形確認: 選択時リボンが「図形の書式」であること",
            "```",
        ]
    )


def session1b_p2(s: dict, pid: int) -> str:
    p4 = s["taskParams"][f"task{pid}_4"]
    p5 = s["taskParams"][f"task{pid}_5"]
    shape_type = p4.get("shapeType", "図形")
    insert_path = P2_SHAPE_INSERT_HINTS.get(shape_type, f"挿入 → 図形 → {shape_type}")
    anchor = p4.get("anchorHeading", "")
    loc_word = "横" if p4.get("shapeLocation") == "besideHeading" else "下"
    return "\n".join(
        [
            "```text",
            "添付の Word 文書を開いてください。2-4用の図形が画像になっている場合、次を実行します。",
            "",
            "1. 画像（図の形式タブになるオブジェクト）を削除",
            f"2. {insert_path} で図形を挿入",
            f"3. 見出し「{anchor}」の{loc_word}に配置（文字列の折り返し: {'四角形' if loc_word == '横' else '行内'}）",
            "4. 図形内は空、または仮の1文字のみ",
            f"5. 「{p4.get('inputText', '')}」は入れない（2-4は受験者が入力）",
            f"6. {p5.get('fontSizePt')}pt・{p5.get('fontStyle')}は設定しない（2-5は受験者が設定）",
            "",
            "【確認】図形クリックで「図形の書式」タブが出ること。",
            "【禁止】挿入→画像、生成画像、最下部テキストボックスのみ",
            f"【保存】{s['workbook']}",
            "```",
        ]
    )


def session2_p2(s: dict, pid: int) -> str:
    return _session2_shell(
        s, pid,
        [
            "| V1 | 2-1 | 切り取り元が本文段落（こちら↓なし）・貼り付け先・リンク段落が設計どおり |",
            "| V2 | 2-2 | 指定見出し下に3項目の箇条書き（初期レベル1、目標レベルは3以外） |",
            "| V3 | 2-3 | 対象文字列が存在し指定色は未設定 |",
            "| V4 | 2-4 | **図形**が指定位置にあり「図形の書式」で編集可（画像・図の形式ならFAIL） |",
            "| V5 | 2-5 | 2-4と同一図形、図形内が空または仮文、指定サイズ・書式は未設定 |",
            "| V6 | 画像禁止 | 装飾画像・2-4代用のPictureが無い |",
            "| V7 | 禁止文字列 | 朗読会・青空文庫・こちら↓関連なし |",
            "| V8 | レイアウト | 余白・向き・スタイルセットが一致 |",
        ],
    )


def session3_p2(pid: int) -> str:
    return "\n".join(
        [
            "```text",
            f"あなたは MOS Word 365 演習アプリの問題文ライターです。",
            f"精査済みパラメータ表に基づき、タスク2-1〜2-5の問題文5件だけを出力してください。",
            "",
            "【文体ルール — 操作種別は教材準拠、パラメータは設計値に従う】",
            "- 2-1: 「○○」の文字列を切り取って、見出し「○○」の下の段落に貼り付けます。（こちら↓形式は使わない）",
            "- 2-2: 見出し「○○」の下にある箇条書きのレベルを「N」に変更します。（Nは設計値、3以外）",
            "- 2-3: 文書内の「○○」の文字の色を「○○」に変更します。（許容色リストから設計値を使用）",
            "- 2-4: 見出し「○○」の下（または横）にある○○の図形に、「\"○○\"」と入力します。",
            "- 2-5: ○○の図形内の文字列を○ptの○○にします。（太字/下線/斜体のいずれか）",
            "",
            "【出力形式】タスク2-1　...（5件）",
            "```",
        ]
    )


# ---------------------------------------------------------------------------
# Project 3 handler
# ---------------------------------------------------------------------------

def variation_table_p3(data: dict, pid: int) -> str:
    rows = []
    for s in data["sets"]:
        p1 = s["taskParams"][f"task{pid}_1"]
        p2 = s["taskParams"][f"task{pid}_2"]
        p3 = s["taskParams"][f"task{pid}_3"]
        p4 = s["taskParams"][f"task{pid}_4"]
        p6 = s["taskParams"][f"task{pid}_6"]
        col = p4.get("columnPreset") or str(p4.get("columns", ""))
        rows.append(
            f"| {s['setNo']} {s['theme']} | {p1['marginPreset']} | {p2.get('breakTypeUi', p2.get('breakType'))} | "
            f"{p3['sectionHeading']}→{p3['orientation']} | {col} | {p6.get('lineSpacingUi', p6.get('lineSpacing'))} |"
        )
    return "\n".join(
        [
            "**タスク3-1〜3-6 のパラメータ（セットごとに異なる）:**",
            "",
            "| セット | 3-1余白 | 3-2区切り | 3-3向き（本文セクション） | 3-4段組 | 3-6行間 |",
            "|--------|--------|----------|------------------------|--------|--------|",
            *rows,
            "",
            "**注意:** 3-1は「やや狭い」以外。3-2は「次のページから」以外。",
            "3-3は文献見出しではなく直前の本文見出しのセクションを指定（縦・横混在）。",
            "3-4は2段以外。3-6は1.3以外。",
            "",
        ]
    )


def session1_p3(s: dict, pid: int, forbidden: list[str]) -> str:
    ly = s["layout"]
    p1 = s["taskParams"][f"task{pid}_1"]
    p2 = s["taskParams"][f"task{pid}_2"]
    p3 = s["taskParams"][f"task{pid}_3"]
    p4 = s["taskParams"][f"task{pid}_4"]
    p5 = s["taskParams"][f"task{pid}_5"]
    p6 = s["taskParams"][f"task{pid}_6"]
    headings = " / ".join(ly.get("headings", []))
    forbidden_line = "、".join(forbidden[:5]) + " 等"
    bib_count = ly.get("bibliographyItemCount", 6)
    col = p4.get("columnPreset") or str(p4.get("columns", ""))
    task_blocks = [
        "1. **3-1 余白**",
        f"   - 初期余白: {ly.get('initialMarginPreset', ly.get('marginPreset'))}（受験者が「{p1['marginPreset']}」に変更）",
        "",
        "2. **3-2〜3-6 文献セクション**",
        f"   - 本文見出し・段落{ly.get('paragraphCount')}前後＋文献見出し「{p2['sectionHeading']}」と文献リスト{bib_count}項目",
        f"   - セクション区切り（{p2.get('breakTypeUi', p2.get('breakType'))}）・向き・段組・段区切り・行間は**未設定**",
        f"   - 3-3対象セクション見出し「{p3['sectionHeading']}」は本文内に配置（向き{p3['orientation']}は未設定）",
        f"   - 3-4段組（{col}）・3-6行間（{p6.get('lineSpacingUi', p6.get('lineSpacing'))}）は未設定",
        f"   - 段区切り対象: 「{p5['columnBreakBefore']}」",
        "",
        "【注意】3-2のセクション区切り挿入後、文献見出しが新セクション先頭になるよう配置すること。",
        "3-3は文献セクションではなく、直前の本文見出しセクションに向きを設定する問題です。",
    ]
    return _session1_shell(s, pid, forbidden_line, headings, ly, task_blocks)


def session2_p3(s: dict, pid: int) -> str:
    p1 = s["taskParams"][f"task{pid}_1"]
    p2 = s["taskParams"][f"task{pid}_2"]
    p3 = s["taskParams"][f"task{pid}_3"]
    p4 = s["taskParams"][f"task{pid}_4"]
    p6 = s["taskParams"][f"task{pid}_6"]
    col = p4.get("columnPreset") or str(p4.get("columns", ""))
    return _session2_shell(
        s, pid,
        [
            f"| V1 | 3-1 | 初期余白が設計どおり（目標「{p1['marginPreset']}」は未設定） |",
            f"| V2 | 3-2 | 文献見出し「{p2['sectionHeading']}」が存在、区切り（{p2.get('breakTypeUi', '')}）未挿入 |",
            f"| V3 | 文献リスト | 項目数・段区切り対象文が設計どおり |",
            f"| V4 | 3-3 | 「{p3['sectionHeading']}」セクションが存在、向き{p3['orientation']}は未設定 |",
            f"| V5 | 3-4〜6 | 段組（{col}）・行間（{p6.get('lineSpacingUi', '')}）が未設定 |",
            "| V6 | 禁止文字列 | 参考文献一覧・五浦美術館関連なし |",
        ],
    )


def session3_p3(pid: int) -> str:
    return "\n".join(
        [
            "```text",
            "精査済みパラメータ表に基づき、タスク3-1〜3-6の問題文6件だけを出力してください。",
            "",
            "【文体ルール — 操作種別は教材準拠、パラメータは設計値に従う】",
            "- 3-1: 文書の余白を「○○」に設定します。（やや狭いは使用しない）",
            "- 3-2: 見出し「○○（文献見出し）」の先頭にセクション区切りを挿入し、○○で始まるように設定します。",
            "      （連続／偶数ページから／奇数ページから。「次のページから」は使わない）",
            "- 3-3: 見出し「○○（文献の直前の本文見出し）」で始まるセクションの印刷の向きを、「縦向き」または「横向き」に設定します。",
            "- 3-4: 見出し「○○（文献見出し）」で始まるセクションを○段組みに設定します。",
            "      （3段／やや狭い2段組み／左段組み／右段組み。標準2段は使わない）",
            "- 3-5: 見出し「○○」の「●…」の先頭に段区切りを挿入します。",
            "- 3-6: 見出し「○○」で始まるセクションの1段目の行間をすべて\"○○\"に設定します。",
            "      （1.5／2／1.15／固定値 18ポイント／1.6行等。1.3は使わない）",
            "",
            "【出力形式】タスク3-1　...（6件）",
            "```",
        ]
    )


# ---------------------------------------------------------------------------
# Project 4 handler
# ---------------------------------------------------------------------------

def _p4_task1_prefill_lines(p1: dict, ly: dict) -> list[str]:
    anchor = p1.get("anchorType", "heading")
    if anchor == "heading":
        return [f"   - 見出し「{p1['heading']}」を配置（コメント**未挿入**）"]
    if anchor == "bodyText":
        return [f"   - 本文に「{p1['targetText']}」を含む段落を配置（コメント**未挿入**）"]
    if anchor == "paragraphUnderHeading":
        return [
            f"   - 見出し「{p1['heading']}」の直下に段落「{p1['targetParagraph']}」を配置",
            "   - コメントは**未挿入**",
        ]
    if anchor == "imageLeftParagraph":
        img = p1.get("imageName", ly.get("commentAnchorImage", ""))
        para = p1.get("targetParagraph", ly.get("commentAnchorParagraph", ""))
        return [
            f"   - 段落「{para}」の右に画像「{img}」を配置（折り返し: 四角形）",
            "   - コメントは段落側に**未挿入**（受験者は左の段落に挿入）",
        ]
    if anchor == "image":
        return [
            f"   - 見出し「{p1['heading']}」の直下に画像「{p1['imageName']}」を配置",
            "   - 画像へのコメントは**未挿入**",
        ]
    return ["   - 4-1 挿入先を設計どおり配置（コメント未挿入）"]


def variation_table_p4(data: dict, pid: int) -> str:
    anchor_labels = {
        "heading": "見出し",
        "bodyText": "本文",
        "paragraphUnderHeading": "見出し下段落",
        "imageLeftParagraph": "画像左段落",
        "image": "画像",
    }
    rows = []
    for s in data["sets"]:
        p1 = s["taskParams"][f"task{pid}_1"]
        p2 = s["taskParams"][f"task{pid}_2"]
        p3 = s["taskParams"][f"task{pid}_3"]
        p4 = s["taskParams"][f"task{pid}_4"]
        p6 = s["taskParams"][f"task{pid}_6"]
        p7 = s["taskParams"][f"task{pid}_7"]
        op = p3.get("operation", "")
        insp = "・".join(p7.get("inspectorRemove", []))
        anchor = anchor_labels.get(p1.get("anchorType", "heading"), p1.get("anchorType", ""))
        rows.append(
            f"| {s['setNo']} {s['theme']} | {anchor} | {p1['commentText']} | {p2['replyText']} | {op} | "
            f"{p4['styleSet']} | {p6['borderWidth']} | {insp} |"
        )
    return "\n".join(
        [
            "**タスク4-1〜4-7 のパラメータ（セットごとに異なる）:**",
            "",
            "| セット | 4-1挿入箇所 | 4-1コメント | 4-2返信 | 4-3操作 | 4-4スタイル | 4-6太さ | 4-7検査削除 |",
            "|--------|------------|-----------|--------|--------|-----------|--------|------------|",
            *rows,
            "",
            "**注意:** 4-1挿入箇所はセットごとに種別が異なる（見出し／本文／見出し下段落／画像左段落／画像）。",
            "4-1に「確認」不可。4-2に「最終確認」不可。4-4は線（シンプル）以外。",
            "4-3は解決のみ／削除のみ／複数解決削除を含む。4-7は透かし+ヘッダーフッター一括以外。",
            "",
        ]
    )


def _p4_task3_prefill_line(p3: dict) -> str:
    op = p3.get("operation", "resolveAndDelete")
    if op == "resolveAndDeleteMultiple":
        texts = "」「".join(p3.get("targetTexts", []))
        return f"   - 本文「{texts}」に未解決コメントを**それぞれ事前挿入**"
    return f"   - 本文「{p3.get('targetText', '')}」に未解決コメントを**事前挿入**"


def _p4_task7_prefill_line(p7: dict, ly: dict) -> str:
    items = p7.get("inspectorRemove", [])
    lines = [f"   - 検査で削除対象: {', '.join(items)}（他項目は残す）"]
    if "ドキュメントのプロパティと個人情報" in items:
        lines.append("   - 文書プロパティ（タイトル・作成者等）を**事前設定**")
    if "非表示テキスト" in items:
        lines.append(f"   - 非表示テキスト「{ly.get('prefilledHiddenText', '—')}」を**事前挿入**")
    if "ヘッダー・フッター" in items and "透かし文字" not in items:
        lines.append(
            f"   - ヘッダー「{ly.get('prefilledHeader', '—')}」・フッター「{ly.get('prefilledFooter', '—')}」を**事前配置**"
        )
    if "コメント" in items:
        lines.append("   - 4-1/4-2/4-3用コメントは事前挿入済み（検査後も残る想定で追加コメント可）")
    return "\n".join(lines)


def session1_p4(s: dict, pid: int, forbidden: list[str]) -> str:
    ly = s["layout"]
    p1 = s["taskParams"][f"task{pid}_1"]
    p2 = s["taskParams"][f"task{pid}_2"]
    p3 = s["taskParams"][f"task{pid}_3"]
    p4 = s["taskParams"][f"task{pid}_4"]
    p5 = s["taskParams"][f"task{pid}_5"]
    p7 = s["taskParams"][f"task{pid}_7"]
    headings = " / ".join(ly.get("headings", []))
    forbidden_line = "、".join(forbidden[:5]) + " 等"
    task_blocks = [
        "1. **4-1 コメント挿入**",
        *_p4_task1_prefill_lines(p1, ly),
        "",
        "2. **4-2 コメント返信**",
        f"   - 見出し「{p2['heading']}」にコメント「{p2.get('prefilledComment', '承認待ち')}」を**事前挿入**",
        f"   - 返信「{p2['replyText']}」は未挿入",
        "",
        "3. **4-3 コメント操作**",
        _p4_task3_prefill_line(p3),
        f"   - 操作種別: {p3.get('operation')}（受験者が実施）",
        "",
        "4. **4-4 スタイルセット**",
        f"   - スタイルセット「{p4['styleSet']}」は**未適用**",
        "",
        "5. **4-5 透かし**",
        f"   - 透かし「{p5['watermarkText']}」を**事前挿入**",
        "",
        "6. **4-6 / 4-7 罫線・検査**",
        "   - ページ罫線は**未設定**",
        _p4_task7_prefill_line(p7, ly),
    ]
    return _session1_shell(s, pid, forbidden_line, headings, ly, task_blocks)


def session2_p4(s: dict, pid: int) -> str:
    return _session2_shell(
        s, pid,
        [
            "| V1 | 4-1 | 挿入先が設計どおり存在、コメント未挿入 |",
            "| V2 | 4-2 | 指定見出しに事前コメントあり、返信なし |",
            "| V3 | 4-3 | 対象文に未解決コメントあり |",
            "| V4 | 4-4 | スタイルセット未適用 |",
            "| V5 | 4-5/7 | 透かし・H/F事前配置済み |",
            "| V6 | 4-6 | ページ罫線未設定 |",
            "| V7 | 禁止文字列 | エコ活動・サンプル２関連なし |",
        ],
    )


def session3_p4(pid: int) -> str:
    return "\n".join(
        [
            "```text",
            "精査済みパラメータ表に基づき、タスク4-1〜4-7の問題文7件だけを出力してください。",
            "",
            "【文体ルール — 操作種別は教材準拠、パラメータは設計値に従う】",
            "- 4-1: 設計値の anchorType に応じて次のいずれか（「確認」は使わない）:",
            "      ・heading: 見出し「○○」に「\"○○\"」とコメントを挿入します。",
            "      ・bodyText: 本文内の「○○」に「\"○○\"」とコメントを挿入します。",
            "      ・paragraphUnderHeading: 見出し「○○」の下の段落「○○」に「\"○○\"」とコメントを挿入します。",
            "      ・imageLeftParagraph: 画像「○○」の左の段落に「\"○○\"」とコメントを挿入します。",
            "      ・image: 見出し「○○」の下の画像「○○」に「\"○○\"」とコメントを挿入します。",
            "- 4-2: 見出し「○○」に挿入されているコメントに「\"○○\"」と返信します。（「最終確認」は使わない）",
            "- 4-3: 設計値の operation に応じて次のいずれか:",
            "      ・resolveOnly: …「○○」に対するコメントを解決してください。",
            "      ・deleteOnly: …「○○」に対するコメントを削除します。",
            "      ・resolveAndDelete: …解決してください。その後、コメントを削除します。",
            "      ・resolveAndDeleteMultiple: …「○○」と「○○」に対するコメントを解決してください。その後、コメントを削除します。",
            "- 4-4: 文書にスタイルセット「○○」を設定します。（線（シンプル）は使わない）",
            "- 4-5: ページの背景に透かし文字「○○」を挿入します。",
            "- 4-6: 文書の周囲を罫線で囲みます。線の色「○○」、線の太さを「○○」に設定します。",
            "- 4-7: ドキュメント検査を行い、○○を削除します。上記以外の項目は削除しません。",
            "      （透かし文字とヘッダー・フッターの一括指定は使わない）",
            "",
            "【出力形式】タスク4-1　...（7件）",
            "```",
        ]
    )


# ---------------------------------------------------------------------------
# Project 5 handler
# ---------------------------------------------------------------------------

def _p5_primary_image_ref(p1: dict) -> str:
    anchor = p1.get("anchorType", "paragraphStart")
    para = p1.get("targetParagraph", "")
    if anchor == "paragraphStart":
        return f"段落「{para}」の先頭の画像"
    if anchor == "paragraphEnd":
        return f"段落「{para}」の末尾の画像"
    if anchor == "underHeading":
        return f"見出し「{p1.get('anchorHeading', '')}」の下の画像"
    if anchor == "besideParagraph":
        return f"見出し「{p1.get('anchorHeading', '')}」の下の段落「{para}」の右の画像"
    if anchor == "underTitle":
        return f"タイトル「{p1.get('titleText', '')}」の下の画像"
    return "5-1で挿入した画像"


def _p5_image_anchor_label(p: dict, title: str = "") -> str:
    anchor = p.get("imageAnchor", "titleRight")
    if anchor == "titleLeft":
        return f"タイトル「{title or p.get('titleText', '')}」の左側の画像"
    if anchor == "titleRight":
        return f"タイトル「{title or p.get('titleText', '')}」の右側の画像"
    if anchor == "underTitle":
        return f"タイトル「{title or p.get('titleText', '')}」の下の画像"
    if anchor == "besideTitle":
        return f"タイトル「{title or p.get('titleText', '')}」の横の画像"
    if anchor == "headingRight":
        return f"見出し「{p.get('anchorHeading', '')}」の右側の画像"
    if anchor == "underHeading":
        return f"見出し「{p.get('anchorHeading', '')}」の下の画像"
    return "指定位置の画像"


def variation_table_p5(data: dict, pid: int) -> str:
    rows = []
    for s in data["sets"]:
        p1 = s["taskParams"][f"task{pid}_1"]
        p2 = s["taskParams"][f"task{pid}_2"]
        p5 = s["taskParams"][f"task{pid}_5"]
        p6 = s["taskParams"][f"task{pid}_6"]
        p7 = s["taskParams"][f"task{pid}_7"]
        p8 = s["taskParams"][f"task{pid}_8"]
        rows.append(
            f"| {s['setNo']} {s['theme']} | {p1.get('anchorType')} / {p1.get('wrapType')} | "
            f"{p2.get('wrapType')} | {p5.get('imageAnchor')} | "
            f"{'説明' if not p6.get('useAltTextWord', True) else '代替テキスト'} | "
            f"{p7.get('wording', '装飾化')} | {p8.get('operation', 'removeBackgroundOnly')} |"
        )
    return "\n".join(
        [
            "**タスク5-1〜5-8 のパラメータ（セットごとに異なる）:**",
            "",
            "| セット | 5-1挿入/折返 | 5-2折返 | 5-5位置 | 5-6表現 | 5-7表現 | 5-8操作 |",
            "|--------|-------------|--------|--------|--------|--------|--------|",
            *rows,
            "",
            "**注意:** 5-1折返しは上下以外。5-2は狭く以外。5-3は水彩：スポンジ以外。",
            "5-4は25pt以外。5-5はハードエッジ以外。5-6は3セット以上で「代替テキスト」不使用。",
            "5-7は3セット以上でスクリーンリーダー非表示表現。5-8は3セット以上で前景領域追加。",
            "",
        ]
    )


def session1_p5(s: dict, pid: int, forbidden: list[str]) -> str:
    ly = s["layout"]
    p1 = s["taskParams"][f"task{pid}_1"]
    p5 = s["taskParams"][f"task{pid}_5"]
    p6 = s["taskParams"][f"task{pid}_6"]
    p8 = s["taskParams"][f"task{pid}_8"]
    headings = " / ".join(ly.get("headings", []))
    forbidden_line = "、".join(forbidden[:5]) + " 等"
    if p1.get("anchorType") == "underHeading":
        p1_line = f"   - 見出し「{p1['anchorHeading']}」の下（画像「{p1['imageName']}」**未挿入**、折り返し{p1['wrapType']}は未設定）"
    elif p1.get("anchorType") == "besideParagraph":
        p1_line = (
            f"   - 見出し「{p1['anchorHeading']}」の下に段落「{p1['targetParagraph']}」を配置し、"
            f"その右に画像「{p1['imageName']}」**未挿入**"
        )
    elif p1.get("anchorType") == "underTitle":
        p1_line = f"   - タイトル「{p1['titleText']}」の下（画像「{p1['imageName']}」**未挿入**）"
    elif p1.get("anchorType") == "paragraphEnd":
        p1_line = f"   - 段落「{p1['targetParagraph']}」の末尾（画像**未挿入**）"
    else:
        p1_line = f"   - 段落「{p1['targetParagraph']}」の先頭（画像「{p1['imageName']}」**未挿入**）"
    p5_line = f"   - {_p5_image_anchor_label(p5)}「{p5['imageName']}」を**事前配置**（面取り未設定）"
    if p6.get("sameImageAs"):
        p6_line = f"   - 5-6対象: {_p5_image_anchor_label(p5, p5.get('titleText', ''))}（説明・代替テキスト**未設定**）"
    else:
        p6_line = f"   - 5-6対象: {_p5_image_anchor_label(p6)}「{p6['imageName']}」を**事前配置**"
    op8 = p8.get("operation", "removeBackgroundOnly")
    p8_note = "前景マーク追加＋背景削除" if op8 == "addForegroundMark" else "背景削除のみ"
    task_blocks = [
        "1. **5-1〜5-4 / 5-7 主画像**",
        p1_line,
        f"   - 5-2〜5-4・5-7の効果はすべて**未設定**",
        "",
        "2. **5-5 / 5-6 副画像**",
        p5_line,
        p6_line,
        "",
        "3. **5-8 文末画像**",
        f"   - 文末に画像「{p8['imageName']}」を**事前配置**（{p8_note}は未実施）",
    ]
    return _session1_shell(s, pid, forbidden_line, headings, ly, task_blocks)


def session2_p5(s: dict, pid: int) -> str:
    return _session2_shell(
        s, pid,
        [
            "| V1 | 5-1 | 挿入先・段落／見出しが設計どおり、主画像未挿入 |",
            "| V2 | 5-5/6 | 副画像の位置・名称が設計どおり、効果・説明未設定 |",
            "| V3 | 5-8 | 文末画像事前配置、背景処理未実施 |",
            "| V4 | 画像名 | 3画像の名称が設計どおり |",
            "| V5 | 禁止文字列 | TOEIC・5月21日関連なし |",
        ],
    )


def session3_p5(pid: int) -> str:
    return "\n".join(
        [
            "```text",
            "精査済みパラメータ表に基づき、タスク5-1〜5-8の問題文8件だけを出力してください。",
            "",
            "【文体ルール — 操作種別は教材準拠、パラメータは設計値に従う】",
            "- 5-1: 設計値の anchorType に応じて挿入（段落先頭／末尾／見出し下）。折り返しは「上下」以外。",
            "- 5-2: 5-1と同一画像の折り返しを変更。「狭く」以外。",
            "- 5-3: 5-1と同一画像にアート効果を設定。「水彩：スポンジ」以外。",
            "- 5-4: 5-1と同一画像にぼかしを設定。「25ポイント」以外。",
            "- 5-5: 設計値の imageAnchor（タイトル左／下／横、見出し右等）の画像に面取りを設定。「ハードエッジ」以外。",
            "- 5-6: 設計値の位置の画像に説明を設定。useAltTextWord=false なら「説明」または「スクリーンリーダー用の説明」（「代替テキスト」という語は使わない）。",
            "- 5-7: wording=装飾化 →「代替テキストを装飾化」／ wording=screenReaderHidden →「スクリーンリーダーに表示されないようにします」",
            "- 5-8: operation=removeBackgroundOnly → 背景削除のみ／ addForegroundMark → 前景領域を一部追加でマーク＋背景以外削除しない",
            "",
            "【出力形式】タスク5-1　...（8件）",
            "```",
        ]
    )


# ---------------------------------------------------------------------------
# Shared session shells
# ---------------------------------------------------------------------------

def _session1_shell(s, pid, forbidden_line, headings, ly, task_blocks):
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 365 演習アプリ向けの類題設計者です。",
            f"project{pid} に相当する演習文書を **1冊** 作成してください。",
            "今回は **Word ファイルの作成とレイアウト仕様の出力** のみ。問題文は書きません。",
            "",
            f"【添付】Word問題文一覧_操作手順.csv（projectId={pid} を参照）",
            "",
            f"【文書名】{s['workbook']}",
            f"【テーマ】{s['theme']}（{s['problemStatement'].replace('あなたは', '').replace('を作成しています。', '')}）",
            "",
            "【レイアウト仕様 — 必ず遵守】",
            layout_table(s),
            "",
            f"【見出し構成】{headings}",
            f"【段落数目安】{ly.get('paragraphCount')}段落（2〜3ページ相当）",
            "",
            "【必須構成とタスク用データ要件】",
            "",
            *task_blocks,
            "",
            "【禁止】",
            f"- 教材の禁止文字列（{forbidden_line}）",
            "- 問題文の生成",
            "- 受験者が行う操作を事前に完了させない",
            "",
            "【出力】",
            "1. docx ファイル",
            "2. Markdown のレイアウト仕様書（見出し一覧・各タスクの段落先頭文・画像名）",
            "```",
        ]
    )


def _session2_shell(s, pid, checklist_rows):
    rows = "\n".join(checklist_rows)
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 類題の検証担当です。",
            f"添付の docx（{s['workbook']}）を開き、以下を検証してください。問題文はまだ書きません。",
            "",
            "【検証対象パラメータ（設計値）】",
            task_params_table(s, pid),
            "",
            "【検証チェックリスト — PASS / FAIL】",
            "",
            "| # | 項目 | 内容 |",
            "|---|------|------|",
            rows,
            "",
            "【出力】精査結果サマリー + 修正済みパラメータ表（実文書の値で確定）",
            "```",
        ]
    )


def session2_p1(s: dict, pid: int) -> str:
    return _session2_shell(
        s, pid,
        [
            "| V1 | 1-1 | 編集記号の初期状態と操作種別が設計どおり |",
            "| V2 | 1-2 | 指定見出し下、開始段落からN段落が未箇条書き |",
            "| V3 | 1-3 | 対象段落・配置・スタイル未設定が設計どおり |",
            "| V4 | 1-4 | 対象段落が変更前スタイルで存在 |",
            "| V5 | 1-5 | 対象段落に事前書式が付与済み |",
            "| V6 | 禁止文字列 | CSR・ゼミ関連文字列なし |",
            "| V7 | レイアウト | 余白・向き・スタイルセット・書式傾向が一致 |",
        ],
    )


def session3_p1(s: dict, pid: int) -> str:
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 365 演習アプリの問題文ライターです。",
            "精査済みパラメータ表に基づき、タスク1-1〜1-5の問題文5件だけを出力してください。",
            "",
            f"【参照】完成初稿は task/MOS_Word類題_project{pid}_配置別_5セット_問題文.md の",
            f"セット{s['setNo']}（{s['theme']}）と整合させること。",
            "",
            "【文体ルール】",
            "- 1-1: 非表示のみ、または表示のみ（設計値に従う）",
            "- 1-2: 見出し「○○」の下の「○○」からNつの段落を箇条書き",
            "- 1-3: 対象段落にスタイル「○○」",
            "- 1-4: 対象段落のスタイルを「○○」に変更",
            "- 1-5: 見出し下の対象段落の書式をすべてクリア",
            "",
            "【出力形式】タスク1-1　...（5件）",
            "```",
        ]
    )


def session3_generic(s: dict, pid: int, rules_fn) -> str:
    body = rules_fn(pid)
    return "\n".join(
        [
            "```text",
            f"【参照】完成初稿は task/MOS_Word類題_project{pid}_配置別_5セット_問題文.md の",
            f"セット{s['setNo']}（{s['theme']}）と整合させること。",
            "",
            body.replace("```text\n", "").replace("\n```", ""),
            "```",
        ]
    )


# ---------------------------------------------------------------------------
# Project 6–10: complete docx helpers
# ---------------------------------------------------------------------------

def _session1_complete_doc_banner() -> list[str]:
    return [
        "【最重要】Microsoft Word 上で docx を直接作成してください（Markdownのみ不可）。",
        "骨組み・プレースホルダ・空表は禁止。見出し1〜4と本文段落で報告書を肉付けし、下記「貼り付け用データ」を**すべて**転記した完全な演習用 docx を1冊作成してください。",
        "受験者が行う操作（表変換・分割・文字効果・置換・変更履歴承認 等）は**未実施**の状態で止めてください。",
        "",
    ]


def usage_section_complete_doc(operation_summary: str) -> list[str]:
    return [
        "## 0. 使い方 — 完全 docx 作成（3セッション × 1セット）",
        "",
        "**Copilot が空の骨組み docx を作るのを防ぐため、セッション1では (1) documentOutline の見出し1〜4＋本文を配置し、(2)「貼り付け用データ」ブロックをそのまま Word に入力してください。**",
        "",
        "```",
        "1チャット = 1セット（§1〜§5 のいずれか1節のみ使用）",
        "      ↓",
        "セッション1 → 該当 § の「セッション1」プロンプト（貼り付け用データを全転記）",
        "      ↓  docx を保存",
        "セッション2 → 同 § の「セッション2」プロンプト（docx 添付）",
        "      ↓",
        "セッション3 → 同 § の「セッション3」プロンプト（検証済みパラメータ表を貼る）",
        "```",
        "",
        "【セッション1で禁止 — ゴミデータ】",
        "- 空の表・空セル・「xxx」「...」「ダミー」「サンプル」",
        "- 行数だけ合わせた中身のない表",
        "- Markdown テキストだけの回答（必ず .docx を保存）",
        "- 受験者が行う操作の事前完了（6-4 文字効果、9-5 承認、10-1 置換 等）",
        "",
        "【セッション1で必須 — 完全類題】",
        "- 貼り付け用データのリテラルを**全セル・全段落に転記**",
        "- テーマに沿った数値（売上・件数・出荷数等）をセットごとに異なる値で記入",
        f"- **操作の大分類を変更しない**（{operation_summary}）",
        "",
        "---",
        "",
    ]


def _render_table_markdown(headers: list, rows: list) -> list[str]:
    lines = ["| " + " | ".join(headers) + " |", "| " + " | ".join(["---"] * len(headers)) + " |"]
    for row in rows:
        lines.append("| " + " | ".join(str(c) for c in row) + " |")
    return lines


def _p6_split_cell_text(p3: dict) -> str:
    return f"{p3['splitRow']}行{p3['splitCol']}列目を{p3['splitInto']}列に"


def _p6_task65_text(p5: dict, heading: str) -> str:
    if p5.get("operation") == "setColumnWidthsAndRowHeight":
        widths = "」「".join(f'"{w}cm"' for w in p5["columnWidthsCm"])
        return (
            f"見出し「{heading}」の表の列幅を左から「{widths}」に設定します。"
            f"行の高さをすべて「\"{p5['rowHeightCm']}cm\"」にします。"
        )
    return f"見出し「{heading}」の表の列幅をすべて同じにします。"


def _render_document_outline_block(s: dict) -> list[str]:
    outline = s.get("contentBlocks", {}).get("documentOutline")
    if not outline:
        return []
    lines = [
        "■ 報告書の肉付け（見出し1〜4・本文段落 — **必須**）",
        "  MOS本番相当の文書量にするため、以下の構成どおりに見出しスタイル1〜4と本文を配置してください。",
        "  タスク用データだけの薄い docx は不可です。",
        "",
    ]
    for line in render_outline_md(outline):
        lines.append(f"  {line}" if line else "")
    lines.append("")
    return lines


def _render_content_blocks_p6(s: dict) -> list[str]:
    cb = s["contentBlocks"]
    p1 = s["taskParams"]["task6_1"]
    p4 = s["taskParams"]["task6_4"]
    p5 = s["taskParams"]["task6_5"]
    p7 = s["taskParams"]["task6_7"]
    tbl = cb["task6_4_table"]
    page = cb.get("task6_7_trailingBlankPage", p7["page"])
    lines = _render_document_outline_block(s)
    if cb.get("introParagraphs"):
        lines.append("■ 導入段落:")
        for p in cb["introParagraphs"]:
            lines.append(f"  - {p}")
    lines += [
        f"■ 貼り付け用データ — 6-1（見出し「{p1['heading']}」の下、タブ区切り{p1['rowCount']}行{p1['colCount']}列・**未表化**）",
    ]
    for row in cb["task6_1_tabRows"]:
        lines.append("\t".join(row))
    lines += [
        "",
        f"■ 貼り付け用データ — 6-4 表「{p4['heading']}」（列幅不均等・1行目装飾なし）",
        *_render_table_markdown(tbl["headers"], tbl["rows"]),
        "",
    ]
    if p5.get("operation") == "setColumnWidthsAndRowHeight":
        lines.append(
            f"■ 6-5用: 上記6-4表は列幅・行高が不均等（目標: 列幅 {', '.join(p5['columnWidthsCm'])}cm、行高 {p5['rowHeightCm']}cm・**未設定**）"
        )
    lines.append(
        f"■ {page}ページ目末尾に空段落を配置（6-7用。{p7['rows']}行{p7['cols']}列表は**未作成**）"
    )
    return lines


def _render_content_blocks_p7(s: dict) -> list[str]:
    cb = s["contentBlocks"]
    p2 = s["taskParams"]["task7_2"]
    lines = _render_document_outline_block(s) + [
        "■ 互換モード docx の作り方（セッション1冒頭）",
        "  1. Word で本文を作成後、「名前を付けて保存」→ ファイルの種類「Word 97-2003」で .doc 保存",
        "  2. 保存した .doc を再度 Word で開く（互換モード表示を確認）",
        f"  3. プロパティ「{p2['propertyName']}」・ヘッダー・フッターは**未設定**のまま",
        "",
        "■ 貼り付け用データ — 導入段落",
    ]
    for p in cb["introParagraphs"]:
        lines.append(f"  - {p}")
    lines.append("■ 貼り付け用データ — 本文段落")
    for p in cb["bodyParagraphs"]:
        lines.append(f"  - {p}")
    return lines


def _p7_preop_task_blocks(s: dict) -> list[str]:
    p2 = s["taskParams"]["task7_2"]
    p3 = s["taskParams"]["task7_3"]
    p4 = s["taskParams"]["task7_4"]
    p6 = s["taskParams"]["task7_6"]
    return [
        "1. **7-1** — 互換モード有効（受験者が解除）",
        f"2. **7-2** — プロパティ「{p2['propertyName']}」**未設定**（会社名は使わない）",
        f"3. **7-3** — ヘッダー「{p3['headerText']}」**未挿入**",
        f"4. **7-4** — 最初のページのヘッダー「{p4['headerText']}」**未設定**",
        "5. **7-5** — フッターのページ番号**未挿入**",
        f"6. **7-6** — フッター「{p6['footerText']}」**未入力**",
        "7. **7-7/7-8** — txt・docm 保存は**未実施**（受験者が実施）",
    ]


P7_OPERATION_SUMMARY = (
    "7-1=互換モード解除、7-2=プロパティ（会社名以外）、7-3=ヘッダー、7-4=最初のページのヘッダー、"
    "7-5=フッターページ番号、7-6=フッターテキスト、7-7=txt保存、7-8=docm保存"
)


def _render_content_blocks_p8(s: dict) -> list[str]:
    cb = s["contentBlocks"]
    p1 = s["taskParams"]["task8_1"]
    ls = cb["linkShape"]
    lines = _render_document_outline_block(s) + [
        f"■ 副題: {cb['subtitle']}",
        f"■ 8-1用: 目次は「{p1.get('placementDescription', '指定位置')}」に**未挿入**（図形内に入れない）",
        "■ 貼り付け用データ — 本文",
    ]
    for heading, paras in cb["bodyUnderHeadings"].items():
        lines.append(f"  見出し「{heading}」:")
        for p in paras:
            lines.append(f"    - {p}")
    lines.append(f"■ 見出し「{s['taskParams']['task8_4']['heading']}」の文献リスト（特殊文字は**未挿入**）")
    for entry in cb["bibliographyEntries"]:
        lines.append(f"  - {entry}")
    lines += [
        f"■ リンク用図形: {ls.get('note', '')}",
        f"  テキスト「{ls['text']}」・見出し「{ls['nearHeading']}」下（リンク**未設定**）",
        "■ 最終ページは空（SmartArt**未挿入**）",
        "■ 脚注は**未挿入**",
    ]
    return lines


def _render_content_blocks_p9(s: dict) -> list[str]:
    cb = s["contentBlocks"]
    ns = cb["numberedSection"]
    t2 = cb["task9_2_table"]
    t3 = cb["task9_3_table"]
    p2 = s["taskParams"]["task9_2"]
    p3 = s["taskParams"]["task9_3"]
    p4 = s["taskParams"]["task9_4"]
    sep = p2.get("separatorLabel", "コンマ区切り")
    side = p3.get("marginSideLabel", "右")
    lines = _render_document_outline_block(s) + [
        f"■ 番号付き段落（見出し「{ns['heading']}」下）",
    ]
    for p in ns["paragraphs"]:
        lines.append(f"  - {p}")
    lines += [
        f"■ 表「{p2['heading']}」（9-2対象・{sep}変換前・表のまま）",
        *_render_table_markdown(t2["headers"], t2["rows"]),
        (
            f"■ 表「{p3['tableHeading']}」（9-3/9-4対象・"
            f"{side}余白{p3['marginMm']}mm未設定・「{p4['sortColumn']}」{p4.get('sortOrderLabel', '降順')}並べ替え未実施）"
        ),
        *_render_table_markdown(t3["headers"], t3["rows"]),
        "■ 変更履歴 ON・以下は**未処理**のまま配置（9-5操作前）",
    ]
    for sample in cb["trackChangesSamples"]:
        lines.append(f"  - [{sample['type']}] {sample['location']}: 「{sample['text']}」")
    return lines


def _render_content_blocks_p10(s: dict) -> list[str]:
    cb = s["contentBlocks"]
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
    if p4["formatType"] == "bold":
        fmt_note = f"「{p4['formatTarget']}」太字未設定"
    elif p4["formatType"] == "color":
        fmt_note = f"「{p4['formatTarget']}」色{p4['formatColor']}未設定"
    else:
        fmt_note = f"「{p4['formatTarget']}」{p4['formatSizePt']}pt未設定"
    lines = _render_document_outline_block(s) + [
        f"■ 貼り付け用データ — 本文（「{p1['searchText']}」が{p1['occurrenceCount']}回以上出現）",
    ]
    for p in cb["bodyWithSearchTerms"]:
        lines.append(f"  - {p}")
    lines.append(f"■ 見出し「{s['taskParams']['task10_2']['heading']}」下の段落（行頭文字・画像は**未設定**）")
    for p in cb["section10_2_paragraphs"]:
        lines.append(f"  - {p}")
    lines.append(
        f"■ 見出し「{p3['heading']}」下の段落"
        f"（{p3['bulletFont']} {p3['bulletCharCode']}は**未設定**）"
    )
    for p in cb["section10_3_paragraphs"]:
        lines.append(f"  - {p}")
    lines.append(f"■ 文書末尾の最後2行（10-5対象・用紙{p5['paperSize']}・行間{ls_label}は**未設定**・{fmt_note}）")
    for p in cb["lastTwoLines"]:
        lines.append(f"  - {p}")
    return lines


CONTENT_BLOCK_RENDERERS = {
    6: _render_content_blocks_p6,
    7: _render_content_blocks_p7,
    8: _render_content_blocks_p8,
    9: _render_content_blocks_p9,
    10: _render_content_blocks_p10,
}


def _session1_complete_shell(s, pid, forbidden_line, headings, ly, task_blocks, paste_blocks):
    extra_layout = []
    if ly.get("pageCount"):
        extra_layout.append(f"| ページ数 | {ly['pageCount']} |")
    if ly.get("compatibilityMode"):
        extra_layout.append("| 互換モード | 有効（受験者が解除） |")
    if ly.get("trackChangesEnabled"):
        extra_layout.append("| 変更履歴 | ON（未承認修正あり） |")
    if ly.get("tablePlacement"):
        extra_layout.append(f"| 表配置 | {ly['tablePlacement']} |")
    layout_extra = "\n".join(extra_layout)

    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 365 演習アプリ向けの類題設計者です。",
            *_session1_complete_doc_banner(),
            f"project{pid} に相当する演習文書を **1冊** 作成します。",
            "問題文テキストは書きません。レイアウトとタスク用データの完全配置が目的です。",
            "",
            f"【添付】Word問題文一覧_操作手順.csv（projectId={pid}）",
            f"【文書名】{s['workbook']}",
            f"【テーマ】{s['theme']}（{s['problemStatement'].replace('あなたは', '').replace('を作成しています。', '')}）",
            "",
            "【レイアウト仕様 — 必ず遵守】",
            layout_table(s),
            layout_extra,
            "",
            f"【見出し構成】{headings}",
            f"【段落数目安】{ly.get('paragraphCount')}段落",
            "",
            "【貼り付け用データ — 以下をすべて Word に転記】",
            "",
            *paste_blocks,
            "",
            "【作成手順 — この順で Word 上で実行】",
            "",
            "■ ステップA: 見出しと本文（肉付け）",
            "  - documentOutline どおりに見出し1〜4と本文段落を配置（タスク用データだけの薄い docx は不可）",
            "  - 文書タイトル・副題（あれば）を配置",
            "  - 上記見出し構成と貼り付け用データを該当位置に転記",
            "",
            "■ ステップB: タスク用データ（下記要件どおり）",
            "",
            *task_blocks,
            "",
            "【禁止】",
            f"  - 教材の禁止文字列（{forbidden_line}）",
            "  - 空セル・ダミー文字・プレースホルダ",
            "  - 問題文の生成",
            "  - 受験者操作の事前完了",
            "",
            "【出力】",
            f"1. ファイル名「{s['workbook']}」で docx を保存",
            "2. 全セル・全段落に具体データが入っていること（骨組み不可）",
            "```",
        ]
    )


def _session2_complete_shell(s, pid, checklist_rows):
    rows = "\n".join(checklist_rows)
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 類題の検証担当です。",
            f"添付の docx（{s['workbook']}）を開き、以下を検証してください。問題文はまだ書きません。",
            "",
            "【検証対象パラメータ（設計値）】",
            task_params_table(s, pid),
            "",
            "【検証チェックリスト — PASS / FAIL】",
            "",
            "| # | 項目 | 内容 |",
            "|---|------|------|",
            "| V_content | 完全性 | 空セル・ダミー文字・プレースホルダが無い |",
            "| V_complete | データ | 貼り付け用データが docx に全て存在（行数・段落数一致） |",
            "| V_preop | 未操作 | 受験者が行う操作が未完了 |",
            rows,
            "",
            "【出力】精査結果サマリー + 修正済みパラメータ表（実文書の値で確定）",
            "```",
        ]
    )


def variation_table_p6(data: dict, pid: int) -> str:
    rows = []
    for s in data["sets"]:
        p1 = s["taskParams"][f"task{pid}_1"]
        p2 = s["taskParams"][f"task{pid}_2"]
        p3 = s["taskParams"][f"task{pid}_3"]
        p5 = s["taskParams"][f"task{pid}_5"]
        p7 = s["taskParams"][f"task{pid}_7"]
        shape = f"{p1['rowCount']}×{p1['colCount']}"
        split3 = f"{p3['splitRow']}行{p3['splitCol']}列→{p3['splitInto']}"
        op5 = "列幅指定+行高" if p5.get("operation") == "setColumnWidthsAndRowHeight" else "列幅均等"
        spec7 = f"p{p7['page']} {p7['rows']}×{p7['cols']}"
        rows.append(
            f"| {s['setNo']} {s['theme']} | {shape} | {p2['splitAtRow']}行目分割 | {split3} | {op5} | {spec7} |"
        )
    return "\n".join(
        [
            f"**タスク{pid}-1〜{pid}-7 のパラメータ（セットごとに異なる）:**",
            "",
            "| セット | 6-1表 | 6-2分割 | 6-3セル分割 | 6-5 | 6-7表 |",
            "|--------|-------|---------|------------|-----|-------|",
            *rows,
            "",
            "**注意:** 6-1は7行2列以外。6-2は5行目以外。6-3は1行2列目→3列以外。6-5は2セットが列幅cm指定＋行高統一。6-7は2ページ13行3列以外。",
            "",
        ]
    )


def session1_p6(s: dict, pid: int, forbidden: list[str]) -> str:
    ly = s["layout"]
    p1 = s["taskParams"][f"task{pid}_1"]
    p2 = s["taskParams"][f"task{pid}_2"]
    p3 = s["taskParams"][f"task{pid}_3"]
    p4 = s["taskParams"][f"task{pid}_4"]
    p5 = s["taskParams"][f"task{pid}_5"]
    p7 = s["taskParams"][f"task{pid}_7"]
    page = s["contentBlocks"].get("task6_7_trailingBlankPage", p7["page"])
    headings = " / ".join(ly.get("headings", []))
    forbidden_line = "、".join(forbidden[:6]) + " 等"
    p5_note = (
        f"列幅 {', '.join(p5['columnWidthsCm'])}cm＋行高 {p5['rowHeightCm']}cm（**未設定**）"
        if p5.get("operation") == "setColumnWidthsAndRowHeight"
        else "列幅均等（**未設定**）"
    )
    task_blocks = [
        f"1. **6-1 表変換** — 見出し「{p1['heading']}」下にタブ{p1['rowCount']}行{p1['colCount']}列（**未表化**）",
        f"2. **6-2 表分割** — 6-1表を{p2['splitAtRow']}行目から分割（**未実施**）",
        f"3. **6-3 セル分割** — 6-1表の{_p6_split_cell_text(p3)}分割（**未実施**）",
        f"4. **6-4 文字効果** — 表「{p4['heading']}」1行目（効果**未設定**、列幅不均等）",
        f"5. **6-5** — 6-4表（{p5_note}）",
        "6. **6-6** — タイトル行は**未設定**",
        f"7. **6-7** — {page}ページ末尾空段落のみ（{p7['rows']}行{p7['cols']}列表は**未作成**）",
    ]
    return _session1_complete_shell(
        s, pid, forbidden_line, headings, ly, task_blocks, _render_content_blocks_p6(s)
    )


def session2_p6(s: dict, pid: int) -> str:
    p1 = s["taskParams"][f"task{pid}_1"]
    p2 = s["taskParams"][f"task{pid}_2"]
    p3 = s["taskParams"][f"task{pid}_3"]
    p5 = s["taskParams"][f"task{pid}_5"]
    p7 = s["taskParams"][f"task{pid}_7"]
    page = s["contentBlocks"].get("task6_7_trailingBlankPage", p7["page"])
    v5 = (
        f"| V5 | 6-5 | 列幅cm指定＋行高{p5.get('rowHeightCm')}cmが未設定 |"
        if p5.get("operation") == "setColumnWidthsAndRowHeight"
        else "| V5 | 6-5 | 列幅不均等（均等化は未実施） |"
    )
    return _session2_complete_shell(
        s, pid,
        [
            f"| V1 | 6-1 | 指定見出し下にタブ{p1['rowCount']}行{p1['colCount']}列テキスト（未表化） |",
            f"| V2 | 6-2/6-3 | 6-1表が存在、{p2['splitAtRow']}行目分割・{_p6_split_cell_text(p3)}分割は未実施 |",
            "| V3 | 6-4 | 第2表が存在、1行目文字効果未設定 |",
            v5,
            f"| V6 | 6-7 | {page}ページ末尾に空段落、{p7['rows']}×{p7['cols']}表は未作成 |",
            "| V7 | 禁止文字列 | 教材表現なし |",
        ],
    )


def session3_p6(pid: int) -> str:
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 365 演習アプリの問題文ライターです。",
            f"精査済みパラメータ表に基づき、タスク6-1〜6-7の問題文7件だけを出力してください。",
            "",
            "【文体ルール — 操作種別は教材準拠、寸法・固有名詞は設計値に従う】",
            "- 6-1: 見出し下の「○○」から「○○」までをN行M列の表に変換（7行2列は使わない）",
            "- 6-2: 表の指定行目から分割（5行目は使わない）",
            "- 6-3: 1つ目の表の指定行・列目を指定列数に分割（1行2列目→3列は使わない）",
            "- 6-4: 表1行目に文字の効果（塗りつぶし＋輪郭）",
            "- 6-5: 列幅をすべて同じに、または列幅cm指定＋行の高さをすべて同じに",
            "- 6-6: 表1行目を各ページのタイトル行に",
            "- 6-7: 指定ページ末尾にN行M列表を作成し1行目にラベル（2ページ13行3列・最高収益等は使わない）",
            "",
            "【出力形式】タスク6-1　...（7件）",
            "```",
        ]
    )


def variation_table_p7(data: dict, pid: int) -> str:
    rows = []
    for s in data["sets"]:
        p2 = s["taskParams"][f"task{pid}_2"]
        p3 = s["taskParams"][f"task{pid}_3"]
        p4 = s["taskParams"][f"task{pid}_4"]
        p6 = s["taskParams"][f"task{pid}_6"]
        p7 = s["taskParams"][f"task{pid}_7"]
        p8 = s["taskParams"][f"task{pid}_8"]
        rows.append(
            f"| {s['setNo']} {s['theme']} | {p2['propertyName']}={p2['propertyValue']} | "
            f"{p3['headerText']} | {p4['headerText']} | {p6['footerText']} | "
            f"{p7['txtBaseName']} | {p8['readPassword']} |"
        )
    return "\n".join(
        [
            f"**タスク{pid}-1〜{pid}-8 のパラメータ:**",
            "",
            "| セット | 7-2プロパティ | 7-3ヘッダー | 7-4先頭頁HDR | 7-6フッター | 7-7保存名 | 7-8PW |",
            "|--------|--------------|------------|-------------|------------|----------|-------|",
            *rows,
            "",
        ]
    )


def usage_section_p7() -> list[str]:
    return [
        "## 0. 使い方",
        "",
        "### 0a. 5セット一括作成（推奨）",
        "",
        "セット1だけ作成してから2〜5を別チャットで作ると品質が落ちやすいため、**初回は §0b の一括プロンプト**を使ってください。",
        "",
        "```",
        "1チャット = セット1〜5を連続作成",
        "      ↓",
        "一括セッション1 → §0b のプロンプト（JSON添付・5ファイル保存）",
        "      ↓",
        "一括セッション2 → §0b 精査プロンプト（5 docx 添付）",
        "      ↓",
        "一括セッション3 → §0b 問題文プロンプト（40件）",
        "```",
        "",
        "### 0b. 1セットずつ作成（§1〜§5）",
        "",
        "**Copilot が空の骨組み docx を作るのを防ぐため、セッション1では (1) documentOutline の見出し1〜4＋本文を配置し、(2)「貼り付け用データ」ブロックをそのまま Word に入力してください。**",
        "",
        "```",
        "1チャット = 1セット（§1〜§5 のいずれか1節のみ使用）",
        "      ↓",
        "セッション1 → 該当 § の「セッション1」プロンプト",
        "      ↓  docx を保存",
        "セッション2 → 同 § の「セッション2」プロンプト（docx 添付）",
        "      ↓",
        "セッション3 → 同 § の「セッション3」プロンプト",
        "```",
        "",
        "【セッション1で禁止 — ゴミデータ】",
        "- 空セル・「xxx」「...」「ダミー」「サンプル」",
        "- Markdown テキストだけの回答（必ず .docx / .doc を保存）",
        "- 受験者操作の事前完了（7-2〜7-6、7-7/7-8の保存）",
        "",
        "【セッション1で必須 — 完全類題】",
        "- documentOutline（見出し1〜4＋本文）を各セットで完全配置",
        "- 導入・本文段落を JSON どおり転記",
        f"- **操作の大分類を変更しない**（{P7_OPERATION_SUMMARY}）",
        "",
        "---",
        "",
    ]


def session1_p7(s: dict, pid: int, forbidden: list[str]) -> str:
    ly = s["layout"]
    headings = " / ".join(ly.get("headings", []))
    forbidden_line = "、".join(forbidden[:4]) + " 等"
    return _session1_complete_shell(
        s, pid, forbidden_line, headings, ly, _p7_preop_task_blocks(s), _render_content_blocks_p7(s)
    )


def session2_p7(s: dict, pid: int) -> str:
    p2 = s["taskParams"]["task7_2"]
    return _session2_complete_shell(
        s, pid,
        [
            "| V1 | 7-1 | 互換モードが有効 |",
            f"| V2 | 7-2 | プロパティ「{p2['propertyName']}」未設定 |",
            "| V3 | 7-3〜7-6 | ヘッダー・フッター未設定 |",
            "| V4 | 本文 | 導入・本文・documentOutline が具体文で存在 |",
            "| V5 | 禁止文字列 | 朗読会・ラビット出版・インテグラル・abc なし |",
        ],
    )


def session3_p7(pid: int) -> str:
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 365 演習アプリの問題文ライターです。",
            f"精査済みパラメータ表に基づき、タスク7-1〜7-8の問題文8件だけを出力してください。",
            "",
            "【文体ルール — 操作種別は教材準拠】",
            "- 7-1: 互換モード解除（OKクリック含む）",
            "- 7-2: プロパティを設定（会社名以外：作成者/件名/管理者/タイトル/分類）",
            "- 7-3: ヘッダーを挿入",
            "- 7-4: 最初のページのみヘッダーに文字入力",
            "- 7-5: フッターにページ番号",
            "- 7-6: フッターにテキスト入力",
            "- 7-7: 名前を付けてテキストファイルとして保存",
            "- 7-8: マクロ有効文書として保存、読み取りパスワード設定",
            "",
            "【出力形式】タスク7-1　...（8件）",
            "```",
        ]
    )


def _p7_batch_spec_table(data: dict, pid: int) -> list[str]:
    rows = []
    for s in data["sets"]:
        ly = s["layout"]
        p2 = s["taskParams"][f"task{pid}_2"]
        p3 = s["taskParams"][f"task{pid}_3"]
        p4 = s["taskParams"][f"task{pid}_4"]
        p6 = s["taskParams"][f"task{pid}_6"]
        p7 = s["taskParams"][f"task{pid}_7"]
        p8 = s["taskParams"][f"task{pid}_8"]
        headings = " / ".join(ly.get("headings", []))
        layout = f"{ly.get('marginPreset')}・{ly.get('styleSet')}・{ly.get('themeColor')}"
        rows.append(
            f"| {s['setNo']} | {s['workbook']} | {headings} | {layout} | "
            f"{p2['propertyName']}={p2['propertyValue']} | {p3['headerText']} | {p4['headerText']} | "
            f"{p6['footerText']} | {p7['txtBaseName']} | {p8['readPassword']} |"
        )
    return [
        "| セット | ファイル名 | 見出し構成 | レイアウト | 7-2 | 7-3 | 7-4 | 7-6 | 7-7 | 7-8PW |",
        "|--------|-----------|-----------|-----------|-----|-----|-----|-----|-----|-------|",
        *rows,
    ]


def session1_batch_p7(data: dict, pid: int, forbidden: list[str]) -> str:
    forbidden_line = "、".join(forbidden[:4]) + " 等"
    set_details = []
    for s in data["sets"]:
        ly = s["layout"]
        headings = " / ".join(ly.get("headings", []))
        set_details += [
            f"--- セット{s['setNo']}（{s['theme']}）: {s['workbook']} ---",
            f"見出し: {headings}",
            f"レイアウト: 余白={ly.get('marginPreset')} / {ly.get('styleSet')} / {ly.get('themeColor')} / {ly.get('format')}",
            "",
            *["  " + line for line in _render_content_blocks_p7(s)],
            "【受験者操作前】",
            *["  " + line for line in _p7_preop_task_blocks(s)],
            "",
        ]
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 365 演習アプリ向けの類題設計者です。",
            *_session1_complete_doc_banner(),
            "",
            "【最重要 — 一括作成】",
            "今回は **セット1〜5の5冊すべて** を、この1回の作業で連続して作成してください。",
            "セット1だけ作って終了しないでください。セット2以降を省略・要約・テンプレ流用しないでください。",
            "各セットは独立した完全な演習用文書として、セット1と **同じ粒度・同じ文書量** で作成してください。",
            "",
            "【作業順序】セット1（カフェ）→ 2（医療）→ 3（製造）→ 4（旅行）→ 5（学習塾）",
            "1セット完了ごとにチェックリストを満たしてから次へ。5冊すべて保存してから最終報告。",
            "",
            f"【添付】MOS_Word類題_project{pid}_配置別_5セット_問題文.json（正本）",
            "【添付】Word問題文一覧_操作手順.csv（projectId=7）",
            "",
            f"【操作種別 — 変更不可】{P7_OPERATION_SUMMARY}",
            "",
            "【5セット仕様一覧】",
            *_p7_batch_spec_table(data, pid),
            "",
            "※ 7-5 は全セット「フッターにページ番号」（未挿入のまま事前配置）",
            "",
            "【各セットの貼り付け用データ・受験者操作前状態】",
            "",
            *set_details,
            "【品質均一化 — セット2〜5を落とさない】",
            "- セット1の見出し階層の深さ・段落数をセット2〜5でも維持",
            "- 「以下同様」「セット1と同構成」で済ませない",
            "- セットごとに intro/body/documentOutline を JSON から個別転記",
            "",
            "【セット完了チェックリスト（各セット保存前）】",
            "□ documentOutline 見出し4まで配置 □ 導入2＋本文4段落 □ 互換モード有効",
            "□ 7-2〜7-6 未操作 □ 禁止文字列なし □ ファイル名一致",
            "",
            "【禁止】",
            f"  - {forbidden_line}",
            "  - プレースホルダ・問題文の生成・受験者操作の事前完了",
            "",
            "【出力】",
            "1. 5ファイルすべて保存したこと",
            "2. 各ファイル名一覧",
            "3. 各セットの7-2〜7-6未操作の確認",
            "```",
        ]
    )


def session2_batch_p7(data: dict, pid: int) -> str:
    files = "、".join(s["workbook"] for s in data["sets"])
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 類題の検証担当です。",
            f"添付の5つの docx（{files}）を **すべて** 検証してください。問題文はまだ書きません。",
            "",
            "【検証方針】セット1を基準品質とし、セット2〜5が同程度の肉付けか比較する。",
            f"JSON（MOS_Word類題_project{pid}_配置別_5セット_問題文.json）の taskParams と照合。",
            "",
            "【各セットで確認】",
            "- 7-1: 互換モード有効",
            "- 7-2〜7-6: 未操作（プロパティ・ヘッダー・フッター未設定）",
            "- intro/body/documentOutline: JSON と一致",
            "- 禁止文字列なし",
            "",
            "【出力】| セット | ファイル名 | PASS/FAIL | 不足点 |（5行）＋ パラメータ表5セット分",
            "```",
        ]
    )


def session3_batch_p7(pid: int) -> str:
    return "\n".join(
        [
            "```text",
            f"【参照】task/類題Json/MOS_Word類題_project{pid}_配置別_5セット_問題文.json の tasks[] と整合させること。",
            "",
            "精査済みの5セットについて、各セット タスク7-1〜7-8 の問題文8件、計40件を出力してください。",
            "セットごとに「## セットN：テーマ」で区切ること。",
            "7-2は会社名以外のプロパティ。操作種別は変更しない。",
            "```",
        ]
    )


def variation_table_p8(data: dict, pid: int) -> str:
    rows = []
    for s in data["sets"]:
        p1 = s["taskParams"][f"task{pid}_1"]
        p2 = s["taskParams"][f"task{pid}_2"]
        p3 = s["taskParams"][f"task{pid}_3"]
        p4 = s["taskParams"][f"task{pid}_4"]
        p5 = s["taskParams"][f"task{pid}_5"]
        p6 = s["taskParams"][f"task{pid}_6"]
        p7 = s["taskParams"][f"task{pid}_7"]
        rows.append(
            f"| {s['setNo']} {s['theme']} | {p1.get('placementDescription', p1.get('tocPlacement'))} | "
            f"{p2['heading']} | {p3['footnoteStartNumber']} | {p4['specialChar']} | "
            f"{p5['smartArtType']} | {p6['smartArtColor']} | {p7['linkType']}/{p7['linkShapeText']} |"
        )
    return "\n".join(
        [
            f"**タスク{pid}-1〜{pid}-7 のパラメータ:**",
            "",
            "| セット | 8-1目次位置 | 8-2見出し | 8-3番号 | 8-4記号 | 8-5SmartArt | 8-6色 | 8-7リンク |",
            "|--------|------------|----------|--------|--------|------------|-------|----------|",
            *rows,
            "",
        ]
    )


def session1_p8(s: dict, pid: int, forbidden: list[str]) -> str:
    ly = s["layout"]
    p1 = s["taskParams"][f"task{pid}_1"]
    p2 = s["taskParams"][f"task{pid}_2"]
    p4 = s["taskParams"][f"task{pid}_4"]
    p7 = s["taskParams"][f"task{pid}_7"]
    headings = " / ".join(ly.get("headings", []))
    forbidden_line = "、".join(forbidden[:6]) + " 等"
    task_blocks = [
        f"1. **8-1** — 「{p1.get('placementDescription')}」に目次**未挿入**（図形内不可）",
        f"2. **8-2** — 「{p2['footnoteAnchor']}」に脚注**未挿入**",
        "3. **8-3** — 脚注番号形式はデフォルト",
        f"4. **8-4** — 見出し「{p4['heading']}」先頭に特殊文字**未挿入**",
        "5. **8-5/8-6** — 最終ページ空、SmartArt**未挿入**",
        f"6. **8-7** — 見出し「{p7['linkNearHeading']}」下の「{p7['linkShapeText']}」図形にリンク**未設定**",
    ]
    return _session1_complete_shell(
        s, pid, forbidden_line, headings, ly, task_blocks, _render_content_blocks_p8(s)
    )


def session2_p8(s: dict, pid: int) -> str:
    p7 = s["taskParams"]["task8_7"]
    return _session2_complete_shell(
        s, pid,
        [
            "| V1 | 8-1 | 指定位置に目次未挿入（図形内・副題下図形なし） |",
            "| V2 | 8-2 | 脚注対象語が本文に存在、脚注未挿入 |",
            "| V3 | 8-4 | 文献見出し先頭に特殊文字未挿入 |",
            "| V4 | 8-5 | 最終ページ空（SmartArt未挿入） |",
            f"| V5 | 8-7 | 「{p7['linkNearHeading']}」下にリンク図形、リンク未設定 |",
            "| V6 | 禁止文字列 | 教材固有語・TOPなし |",
        ],
    )


def session3_p8(pid: int) -> str:
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 365 演習アプリの問題文ライターです。",
            f"精査済みパラメータ表に基づき、タスク8-1〜8-7の問題文7件だけを出力してください。",
            "",
            "【文体ルール — 操作種別は教材準拠・パラメータは設計値】",
            "- 8-1: 図形内・副題下以外の指定位置に目次（自動作成の目次2）",
            "- 8-2: 指定見出し下の指定語の後ろに脚注（って何？見出しは使わない）",
            "- 8-3: 脚注開始番号を①以外（*、壱、ⅰ、※、ア等）に変更",
            "- 8-4: 見出し先頭に§以外の特殊文字（©、™、¶、®、⋯等）",
            "- 8-5: 最終ページに基本ベン図以外のSmartArt＋3語",
            "- 8-6: SmartArt色をカラフル以外（モノクロ、アクセント1〜6）",
            "- 8-7: 最下部以外の図形にリンク（見出し・URL・文末・文頭）",
            "",
            "【出力形式】タスク8-1　...（7件）",
            "```",
        ]
    )


def variation_table_p9(data: dict, pid: int) -> str:
    rows = []
    for s in data["sets"]:
        p1 = s["taskParams"][f"task{pid}_1"]
        p2 = s["taskParams"][f"task{pid}_2"]
        p3 = s["taskParams"][f"task{pid}_3"]
        p4 = s["taskParams"][f"task{pid}_4"]
        p5 = s["taskParams"][f"task{pid}_5"]
        p6 = s["taskParams"][f"task{pid}_6"]
        rows.append(
            f"| {s['setNo']} {s['theme']} | {p1['restartAt']} | {p2.get('separatorLabel', 'コンマ')} | "
            f"{p3.get('marginSideLabel', '右')}{p3['marginMm']}mm | "
            f"{p4['sortColumn']}{p4.get('sortOrderLabel', '降順')} | "
            f"{p5['operation']} | {p6['trackChangesPassword']} |"
        )
    return "\n".join(
        [
            f"**タスク{pid}-1〜{pid}-6 のパラメータ:**",
            "",
            "| セット | 9-1再開 | 9-2区切り | 9-3余白 | 9-4並べ替え | 9-5操作 | パスワード |",
            "|--------|---------|----------|---------|------------|---------|-----------|",
            *rows,
            "",
        ]
    )


def session1_p9(s: dict, pid: int, forbidden: list[str]) -> str:
    ly = s["layout"]
    headings = " / ".join(ly.get("headings", []))
    forbidden_line = "、".join(forbidden[:4]) + " 等"
    p2 = s["taskParams"]["task9_2"]
    p3 = s["taskParams"]["task9_3"]
    p4 = s["taskParams"]["task9_4"]
    sep = p2.get("separatorLabel", "コンマ区切り")
    side = p3.get("marginSideLabel", "右")
    task_blocks = [
        "1. **9-1** — 番号付き段落が配置済み（再開値は未変更）",
        f"2. **9-2** — 対象表は表のまま（{sep}変換**未実施**）",
        f"3. **9-3/9-4** — 対象表の{side}余白{p3['marginMm']}mm**未設定**、"
        f"「{p4['sortColumn']}」{p4.get('sortOrderLabel', '降順')}並べ替え**未実施**",
        "4. **9-5** — 変更履歴ON、未処理の修正あり（9-5の操作**未実施**）",
        "5. **9-6** — 変更履歴ロック**未設定**",
    ]
    return _session1_complete_shell(
        s, pid, forbidden_line, headings, ly, task_blocks, _render_content_blocks_p9(s)
    )


def session2_p9(s: dict, pid: int) -> str:
    return _session2_complete_shell(
        s, pid,
        [
            "| V1 | 9-1 | 番号付き段落と対象段落が存在 |",
            "| V2 | 9-2/9-3 | 2つの表が具体データで存在 |",
            "| V3 | 9-5 | 変更履歴ON、未承認修正2箇所以上 |",
            "| V4 | 9-5前 | 承認・記録停止は未実施 |",
            "| V5 | 禁止文字列 | ゴールデンウィーク・654等なし |",
        ],
    )


def session3_p9(pid: int) -> str:
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 365 演習アプリの問題文ライターです。",
            f"精査済みパラメータ表に基づき、タスク9-1〜9-6の問題文6件だけを出力してください。",
            "",
            "【文体ルール — 操作種別は教材準拠】",
            "- 9-1: 段落番号を指定値（1以外）から再開",
            "- 9-2: 表をコンマ/タブ/段落区切りの文字列に変更（コンマは2セットのみ）",
            "- 9-3: セル余白（右・4mm以外の辺と数値）",
            "- 9-4: 指定列の昇順・降順・あいうえお順並べ替え（合計列降順は不可）",
            "- 9-5: 変更履歴の承認のみ / 修正前に戻す / 取り消し＋記録停止",
            "- 9-6: 変更履歴ロック＋パスワード",
            "",
            "【出力形式】タスク9-1　...（6件）",
            "```",
        ]
    )


def variation_table_p10(data: dict, pid: int) -> str:
    rows = []
    for s in data["sets"]:
        p1 = s["taskParams"][f"task{pid}_1"]
        p2 = s["taskParams"][f"task{pid}_2"]
        p3 = s["taskParams"][f"task{pid}_3"]
        p4 = s["taskParams"][f"task{pid}_4"]
        p5 = s["taskParams"][f"task{pid}_5"]
        ft = p4["formatType"]
        if ft == "bold":
            fmt = f"太字:{p4['formatTarget']}"
        elif ft == "color":
            fmt = f"色{p4['formatColor']}:{p4['formatTarget']}"
        else:
            fmt = f"{p4['formatSizePt']}pt:{p4['formatTarget']}"
        lst = p5["lineSpacingType"]
        val = p5.get("lineSpacingValue")
        ls = lst if lst in ("1行", "1.5行", "2行") else f"{lst}{val or ''}"
        rows.append(
            f"| {s['setNo']} {s['theme']} | {p1['searchText']}→{p1['replaceText']} | "
            f"{p3['bulletFont']}/{p3['bulletCharCode']} | {fmt} | {p5['paperSize']}/{ls} |"
        )
    return "\n".join(
        [
            f"**タスク{pid}-1〜{pid}-5 のパラメータ:**",
            "",
            "| セット | 置換 | 10-3フォント/コード | 10-4書式 | 10-5用紙/行間 |",
            "|--------|------|-------------------|---------|-------------|",
            *rows,
            "",
        ]
    )


def session1_p10(s: dict, pid: int, forbidden: list[str]) -> str:
    ly = s["layout"]
    p1 = s["taskParams"][f"task{pid}_1"]
    p3 = s["taskParams"][f"task{pid}_3"]
    p4 = s["taskParams"][f"task{pid}_4"]
    p5 = s["taskParams"][f"task{pid}_5"]
    headings = " / ".join(ly.get("headings", []))
    forbidden_line = "、".join(forbidden[:6]) + " 等"
    lst = p5["lineSpacingType"]
    val = p5.get("lineSpacingValue")
    ls = lst if lst in ("1行", "1.5行", "2行") else f"{lst}{val or ''}"
    if p4["formatType"] == "bold":
        fmt = f"「{p4['formatTarget']}」太字"
    elif p4["formatType"] == "color":
        fmt = f"「{p4['formatTarget']}」色{p4['formatColor']}"
    else:
        fmt = f"「{p4['formatTarget']}」{p4['formatSizePt']}pt"
    task_blocks = [
        f"1. **10-1** — 「{p1['searchText']}」が本文に{p1['occurrenceCount']}回以上（置換**未実施**）",
        "2. **10-2** — 指定見出し下は通常段落（画像行頭文字**未設定**）",
        f"3. **10-3** — 指定見出し下は通常段落（{p3['bulletFont']} {p3['bulletCharCode']}**未設定**）",
        f"4. **10-4** — {fmt}は**未設定**",
        f"5. **10-5** — 用紙{p5['paperSize']}・最後2行行間{ls}は**未設定**",
    ]
    return _session1_complete_shell(
        s, pid, forbidden_line, headings, ly, task_blocks, _render_content_blocks_p10(s)
    )


def session2_p10(s: dict, pid: int) -> str:
    p1 = s["taskParams"][f"task{pid}_1"]
    return _session2_complete_shell(
        s, pid,
        [
            f"| V1 | 10-1 | 「{p1['searchText']}」が{p1['occurrenceCount']}回以上、置換未実施 |",
            "| V2 | 10-2/10-3 | 見出し下に具体段落が存在、行頭文字未設定 |",
            "| V3 | 10-5 | 文書末尾に最後2行が識別可能 |",
            "| V4 | 禁止文字列 | ウイルス・事例・PCアイコン等なし |",
        ],
    )


def session3_p10(pid: int) -> str:
    return "\n".join(
        [
            "```text",
            "あなたは MOS Word 365 演習アプリの問題文ライターです。",
            f"精査済みパラメータ表に基づき、タスク10-1〜10-5の問題文5件だけを出力してください。",
            "",
            "【文体ルール — 操作種別は教材準拠】",
            "- 10-1: 一括置換（置換前後はセットごとに独立した語句）",
            "- 10-2: 画像の行頭文字に変更",
            "- 10-3: 記号フォントの行頭文字（Webdings・120以外）",
            "- 10-4: 置換で太字 / フォント色 / 文字サイズ（斜体不可）",
            "- 10-5: 用紙サイズ（B5以外）、最後2行の行間（1行/1.5行/2行/最小値/固定値/倍数、1.6行不可）",
            "",
            "【出力形式】タスク10-1　...（5件）",
            "```",
        ]
    )


# ---------------------------------------------------------------------------
# Handler registry
# ---------------------------------------------------------------------------

PROJECT_HANDLERS = {
    1: {
        "variation_table": variation_table_p1,
        "session1": session1_p1,
        "session2": session2_p1,
        "session3": session3_p1,
        "ui_lines": lambda ui: [
            f"- 編集記号: {ui.get('showHideMarks', '—')}",
            f"- 箇条書き: {ui.get('bulletList', '—')}",
            f"- スタイル: {ui.get('styleGallery', '—')}",
            f"- 書式クリア: {ui.get('clearFormatting', '—')}",
        ],
        "operation_summary": "1-1=編集記号、1-2=箇条書き、1-3=スタイル適用、1-4=スタイル変更、1-5=書式クリア",
    },
    2: {
        "variation_table": variation_table_p2,
        "usage_section": usage_section_p2,
        "session1": session1_p2,
        "session1b": session1b_p2,
        "session2": session2_p2,
        "session3": lambda s, pid: session3_generic(s, pid, session3_p2),
        "ui_lines": lambda ui: [
            f"- 切り取り/貼り付け: {ui.get('cutPaste', '—')}",
            f"- 箇条書きレベル: {ui.get('bulletLevel', '—')}",
            f"- フォントの色: {ui.get('fontColor', '—')}",
            f"- 図形挿入: {ui.get('shapeText', '挿入 → 図形')}",
            f"- 図形内書式: {ui.get('fontFormat', '—')}",
        ],
        "operation_summary": "2-1=切り取り貼り付け、2-2=箇条書きレベル、2-3=文字色、2-4=図形入力、2-5=図形内書式",
    },
    3: {
        "variation_table": variation_table_p3,
        "session1": session1_p3,
        "session2": session2_p3,
        "session3": lambda s, pid: session3_generic(s, pid, session3_p3),
        "ui_lines": lambda ui: [
            f"- 余白: {ui.get('pageMargins', '—')}",
            f"- セクション区切り: {ui.get('sectionBreak', '—')}",
            f"- 段組み: {ui.get('columns', '—')}",
        ],
        "operation_summary": "3-1=余白、3-2=セクション区切り、3-3=向き、3-4=段組み、3-5=段区切り、3-6=行間",
    },
    4: {
        "variation_table": variation_table_p4,
        "session1": session1_p4,
        "session2": session2_p4,
        "session3": lambda s, pid: session3_generic(s, pid, session3_p4),
        "ui_lines": lambda ui: [
            f"- コメント: {ui.get('insertComment', '—')}",
            f"- スタイルセット: {ui.get('styleSet', '—')}",
            f"- 透かし: {ui.get('watermark', '—')}",
            f"- ドキュメント検査: {ui.get('documentInspector', '—')}",
        ],
        "operation_summary": "4-1=コメント、4-2=返信、4-3=解決削除、4-4=スタイルセット、4-5=透かし、4-6=罫線、4-7=検査",
    },
    5: {
        "variation_table": variation_table_p5,
        "session1": session1_p5,
        "session2": session2_p5,
        "session3": lambda s, pid: session3_generic(s, pid, session3_p5),
        "ui_lines": lambda ui: [
            f"- 画像挿入: {ui.get('insertPicture', '—')}",
            f"- 折り返し: {ui.get('wrapText', '—')}",
            f"- 代替テキスト: {ui.get('altText', '—')}",
            f"- 背景削除: {ui.get('removeBackground', '—')}",
        ],
        "operation_summary": "5-1=画像挿入、5-2=折り返し、5-3=アート効果、5-4=ぼかし、5-5=面取り、5-6=代替テキスト、5-7=装飾化、5-8=背景削除",
    },
    6: {
        "variation_table": variation_table_p6,
        "usage_section": lambda: usage_section_complete_doc(
            "6-1=表変換、6-2=表分割、6-3=セル分割、6-4=文字効果、6-5=列幅均等または列幅指定＋行高、6-6=タイトル行、6-7=表作成"
        ),
        "session1": session1_p6,
        "session2": session2_p6,
        "session3": lambda s, pid: session3_generic(s, pid, session3_p6),
        "ui_lines": lambda ui: [
            f"- 表変換: {ui.get('convertToTable', '—')}",
            f"- 表分割: {ui.get('splitTable', '—')}",
            f"- 文字効果: {ui.get('textEffect', '—')}",
            f"- 列幅/行高: {ui.get('distributeColumns', '—')} / {ui.get('columnWidth', '—')} / {ui.get('rowHeight', '—')}",
            f"- 表作成: {ui.get('insertTable', '—')}",
        ],
        "operation_summary": "6-1=表変換、6-2=表分割、6-3=セル分割、6-4=文字効果、6-5=列幅均等または列幅指定＋行高、6-6=タイトル行、6-7=表作成",
    },
    7: {
        "variation_table": variation_table_p7,
        "usage_section": usage_section_p7,
        "batch_session1": session1_batch_p7,
        "batch_session2": session2_batch_p7,
        "batch_session3": session3_batch_p7,
        "session1": session1_p7,
        "session2": session2_p7,
        "session3": lambda s, pid: session3_generic(s, pid, session3_p7),
        "ui_lines": lambda ui: [
            f"- 互換モード: {ui.get('compatibilityMode', '—')}",
            f"- プロパティ: {ui.get('documentProperties', '—')}",
            f"- ヘッダー: {ui.get('header', '—')}",
            f"- フッター: {ui.get('footer', '—')}",
            f"- 保存: {ui.get('saveAsTxt', '—')} / {ui.get('saveAsDocm', '—')}",
        ],
        "operation_summary": P7_OPERATION_SUMMARY,
    },
    8: {
        "variation_table": variation_table_p8,
        "usage_section": lambda: usage_section_complete_doc(
            "8-1=目次（図形外）、8-2=脚注、8-3=脚注番号、8-4=特殊文字、"
            "8-5=SmartArt、8-6=SmartArt色、8-7=リンク（最下部以外）"
        ),
        "session1": session1_p8,
        "session2": session2_p8,
        "session3": lambda s, pid: session3_generic(s, pid, session3_p8),
        "ui_lines": lambda ui: [
            f"- 目次: {ui.get('tableOfContents', '—')}",
            f"- 脚注: {ui.get('footnote', '—')}",
            f"- SmartArt: {ui.get('smartArt', '—')}",
            f"- リンク: {ui.get('hyperlink', '—')}",
        ],
        "operation_summary": (
            "8-1=目次（図形外）、8-2=脚注、8-3=脚注番号、8-4=特殊文字、"
            "8-5=SmartArt、8-6=SmartArt色、8-7=リンク（最下部以外）"
        ),
    },
    9: {
        "variation_table": variation_table_p9,
        "usage_section": lambda: usage_section_complete_doc(
            "9-1=番号再開(1以外)、9-2=表→文字列(コンマ/タブ/段落)、9-3=セル余白、9-4=並べ替え、9-5=変更履歴操作、9-6=変更履歴ロック"
        ),
        "session1": session1_p9,
        "session2": session2_p9,
        "session3": lambda s, pid: session3_generic(s, pid, session3_p9),
        "ui_lines": lambda ui: [
            f"- 番号再開: {ui.get('restartNumbering', '—')}",
            f"- 表変換: {ui.get('convertToText', '—')}",
            f"- 変更履歴: {ui.get('trackChanges', '—')}",
        ],
        "operation_summary": "9-1=番号再開、9-2=表→文字列、9-3=セル余白、9-4=並べ替え、9-5=変更履歴承認、9-6=変更履歴ロック",
    },
    10: {
        "variation_table": variation_table_p10,
        "usage_section": lambda: usage_section_complete_doc(
            "10-1=一括置換、10-2=画像行頭文字、10-3=記号フォント行頭文字、10-4=置換書式、10-5=用紙＋行間"
        ),
        "session1": session1_p10,
        "session2": session2_p10,
        "session3": lambda s, pid: session3_generic(s, pid, session3_p10),
        "ui_lines": lambda ui: [
            f"- 置換: {ui.get('replaceAll', '—')}",
            f"- 行頭文字: {ui.get('bulletPicture', '—')}",
            f"- 用紙: {ui.get('pageSize', '—')}",
        ],
        "operation_summary": "10-1=一括置換、10-2=画像行頭文字、10-3=記号フォント行頭文字、10-4=置換書式、10-5=用紙＋行間",
    },
}


def generate(pid: int) -> str:
    data = json.loads(
        (BASE / "類題Json" / f"MOS_Word類題_project{pid}_配置別_5セット_問題文.json").read_text(
            encoding="utf-8"
        )
    )
    handler = PROJECT_HANDLERS[pid]
    forbidden = data.get("variantRules", {}).get("forbiddenParagraphTexts", [])
    task_count = len(data["sets"][0]["tasks"])
    ui = data.get("wordUiNames", {})

    lines = [
        f"# Word 類題作成 — Copilot 用プロンプト（Project{pid} 配置別）",
        "",
        "Microsoft Copilot に **コピペして使う** プロンプト集です。",
        f"project{pid} の類題文書（5セット×{task_count}タスク）を作成します。",
        "",
        "**添付ファイル（推奨）:**",
        f"- `@task/Word問題文一覧_操作手順.csv`（projectId={pid} の行）",
        "",
        "**問題文ファイル（完成初稿）:**",
        f"- [`task/MOS_Word類題_project{pid}_配置別_5セット_問題文.md`](MOS_Word類題_project{pid}_配置別_5セット_問題文.md)",
        f"- [`task/類題Json/MOS_Word類題_project{pid}_配置別_5セット_問題文.json`](類題Json/MOS_Word類題_project{pid}_配置別_5セット_問題文.json)",
        f"- [`task/類題Json/MOS_Word類題_project{pid}_配置別_5セット_問題文_アプリ用.json`](類題Json/MOS_Word類題_project{pid}_配置別_5セット_問題文_アプリ用.json)",
        "",
        f"**禁止文字列（教材流用不可）:** {' '.join(f'`{x}`' for x in forbidden)}",
        "",
        handler["variation_table"](data, pid),
        "**Word UI パス（ビルド2508）:**",
        "",
        *handler["ui_lines"](ui),
        "",
        "---",
        "",
    ]
    if handler.get("usage_section"):
        us = handler["usage_section"]
        lines.extend(us() if callable(us) else us)

    if handler.get("batch_session1"):
        lines += [
            "## §0b 5セット一括作成（推奨）",
            "",
            "セット1のみ別チャットで作成するとセット2〜5の品質が落ちるため、**初回はこちらを使用**してください。",
            "",
            "### 一括セッション1 — セット1〜5を連続作成",
            "",
            handler["batch_session1"](data, pid, forbidden),
            "",
            "### 一括セッション2 — 5ファイル精査",
            "",
            handler["batch_session2"](data, pid),
            "",
            "### 一括セッション3 — 問題文40件",
            "",
            handler["batch_session3"](pid),
            "",
            "---",
            "",
        ]

    if not handler.get("usage_section"):
        lines += [
            "## 0. 使い方（3セッション × 1セット）",
            "",
            "```",
            "1チャット = 1セット（§1〜§5 のいずれか1節のみ使用）",
            "      ↓",
            "セッション1 → 該当 § の「セッション1」プロンプト",
            "      ↓  docx を保存",
            "セッション2 → 同 § の「セッション2」プロンプト（docx 添付）",
            "      ↓",
            "セッション3 → 同 § の「セッション3」プロンプト（検証済みパラメータ表を貼る）",
            "```",
            "",
            "- **問題文はセッション3まで書かない**（セッション1はレイアウト仕様のみ）",
            f"- **操作の大分類を変更しない**（{handler['operation_summary']}）",
            f"- 文書名: `MOS_Word類題_project{pid}_配置別_セット{{N}}_{{テーマ}}.docx`",
            "",
            "---",
            "",
        ]

    for s in data["sets"]:
        n = s["setNo"]
        lines += [
            f"## §{n} セット{n} — {s['theme']}",
            "",
            f"**文書名:** `{s['workbook']}`",
            "",
            "### セッション1 — Word 上で文書を作成（本文＋図形）",
            "",
            handler["session1"](s, pid, forbidden),
            "",
        ]
        if handler.get("session1b"):
            lines += [
                "### セッション1b — 図形の差し替え（画像になった場合のみ）",
                "",
                handler["session1b"](s, pid),
                "",
            ]
        lines += [
            "### セッション2 — 精査",
            "",
            handler["session2"](s, pid),
            "",
            "### セッション3 — 問題文生成",
            "",
            handler["session3"](s, pid),
            "",
            "### 設計値（taskParams）",
            "",
            task_params_table(s, pid),
            "",
            "---",
            "",
        ]

    lines += ["## 完成問題文（初稿）", ""]
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
        f"| 2026-06-26 | Word Project{pid} 配置別5セット初版 |",
        "",
    ]
    return "\n".join(lines)


if __name__ == "__main__":
    for p in sys.argv[1:] or [1]:
        pid = int(p)
        if pid not in PROJECT_HANDLERS:
            print(f"skip unknown project {pid}")
            continue
        out = BASE / f"Word_類題_Copilot_project{pid}_配置別.md"
        out.write_text(generate(pid), encoding="utf-8")
        print(f"wrote {out}")
