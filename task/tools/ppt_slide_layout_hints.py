"""Exam-oriented layout hints for PowerPoint Copilot prompts (P4–P10)."""

from __future__ import annotations


def _shape_by_id(slide: dict, shape_id: str) -> dict | None:
    for sh in slide.get("shapes", []):
        if sh.get("id") == shape_id:
            return sh
    return None


def render_exam_layout_hints(
    project_id: int,
    slide_map: list[dict],
    task_params: dict,
) -> list[str]:
    if project_id == 4:
        return _layout_p4(slide_map, task_params)
    if project_id == 5:
        return _layout_p5(slide_map, task_params)
    if project_id == 6:
        return _layout_p6(slide_map, task_params)
    if project_id == 7:
        return _layout_p7(slide_map, task_params)
    if project_id == 8:
        return _layout_p8(slide_map, task_params)
    if project_id == 9:
        return _layout_p9(slide_map, task_params)
    if project_id == 10:
        return _layout_p10(slide_map, task_params)
    return []


def _header() -> list[str]:
    return [
        "【試験用レイアウト仕様 — 図形・画像の配置を厳守】",
        "  ※ 挿入→画像／挿入→図形／挿入→アイコン を使用。SmartArt・汎用アイコン・家形プレースホルダは禁止。",
        "  ※ 下記の初期ずれ・未設定は受験者タスク用。事前に直さないこと。",
        "",
    ]


def _layout_p4(slide_map: list[dict], task_params: dict) -> list[str]:
    lines = _header()
    p1 = task_params.get("task4_1", {})
    p2 = task_params.get("task4_2", {})

    for s in slide_map:
        no = s["slideNo"]
        title = s["title"]
        lines.append(f"  スライド{no}「{title}」")

        if no == 1:
            sh = _shape_by_id(s, "inputShape")
            img = _shape_by_id(s, "heroImage")
            if sh:
                lines += [
                    f"    - 挿入→図形「{sh['insertName']}」1個: 塗り={sh.get('fillColorDesc', '')}、"
                    f"図形内文字は空のまま（4-1「{p1.get('shapeText', '')}」未入力）",
                ]
            if img:
                lines += [
                    f"    - 挿入→画像1枚: ラベル／代替テキスト「{img.get('label', '')}」、"
                    "スライド右半分に配置",
                    "    - 光彩・クイックスタイル・アート効果は未設定（4-3/4-4未実施）",
                ]
            lines.append("    - 箇条書きはタイトル下プレースホルダに配置（図形と重ねない）")

        elif no == 2:
            sh = _shape_by_id(s, "textBoxShape")
            if sh:
                lines += [
                    f"    - 挿入→図形「{sh['insertName']}」1個をコンテンツ領域に配置",
                    "    - 塗りつぶし・枠線はデフォルトのまま（4-7未実施）",
                ]
            lines.append("    - 箇条書きは図形の左または下のプレースホルダに配置")

        elif no == 3:
            sh = _shape_by_id(s, "centerTextBox")
            if sh:
                lines += [
                    f"    - 挿入→図形「{sh['insertName']}」1個をコンテンツ領域の上寄せに配置",
                    "    - 垂直方向中央揃えは未実施（4-8）。中央ではなく上寄せのまま",
                ]
            lines.append("    - 箇条書きはプレースホルダに配置")

        elif no == 4:
            phrase = p2.get("targetPhrase", "")
            lines += [
                f"    - テキストボックスに「{phrase}」を1行で入力（一字一句一致）",
                f"    - 文字の塗りつぶしは未設定（4-2「{p2.get('fillColor', '')}」未適用）",
                "    - 箇条書きは別プレースホルダに配置（フレーズ文字列と混同しない）",
            ]

        elif no == 5:
            left = _shape_by_id(s, "imageLeft")
            right = _shape_by_id(s, "imageRight")
            left_label = left.get("label", "") if left else ""
            right_label = right.get("label", "") if right else ""
            lines += [
                "    - 挿入→画像の写真2枚（アイコン・図形禁止）。縦2段配置は禁止",
                f"    - 左画像=「{left_label}」: スライド中央やや左、高さは中〜大",
                f"    - 右画像=「{right_label}」: 左画像の右隣に横並び",
                "    - 箇条書きはプレースホルダ左1/3。画像は中央〜右半分（箇条書きと画像を分離）",
                "    - 初期状態: 右画像の上端は左画像より下（上端を揃えない・4-5未実施）",
                "    - 初期状態: 右画像の右端はスライド右端からはみ出す（4-6未実施・トリミング禁止）",
                "    - 画像の位置揃え・トリミングは未実施のまま保存",
            ]

        else:
            lines.append("    - 箇条書きはタイトル下プレースホルダに配置")

        lines.append("")

    return lines


def _layout_p5(slide_map: list[dict], task_params: dict) -> list[str]:
    lines = _header()
    p1 = task_params.get("task5_1", {})
    p3 = task_params.get("task5_3", {})
    p4 = task_params.get("task5_4", {})
    p7 = task_params.get("task5_7", {})

    for s in slide_map:
        no = s["slideNo"]
        title = s["title"]
        lines.append(f"  スライド{no}「{title}」")

        if no == 1:
            lines.append("    - 箇条書きのみ（図形なし）。タイトル下プレースホルダに配置")

        elif no == 2:
            img = _shape_by_id(s, "decorativeImage")
            icon = _shape_by_id(s, "helpIcon")
            if img:
                lines.append(
                    f"    - 挿入→画像1枚「{img.get('label', '')}」: コンテンツ領域左寄り"
                )
            if icon:
                lines.append(
                    f"    - 挿入→アイコン「{icon.get('label', '')}」: 画像の右隣、"
                    f"塗りつぶし未設定（5-7「{p7.get('iconFillColor', '')}」未適用）"
                )
            lines.append("    - 代替テキスト装飾化は未実施（5-6）")
            lines.append("    - 箇条書きはプレースホルダに配置")

        elif no == 3:
            st = p1.get("alignShapeType", "")
            lines += [
                f"    - 挿入→図形「{p1.get('alignShapeInsertName', st)}」を4個、横一列にばらばら配置",
                "    - 4個の右端は揃えない（5-1右揃え未実施）。高さ・間隔は揃っていなくてよい",
                "    - 箇条書きはプレースホルダに配置",
            ]

        elif no == 4:
            lines += [
                f"    - 挿入→図形「{p3.get('sourceShapeInsertName', '')}」1個（{p3.get('sourceShapeType', '')}）",
                f"    - {p3.get('targetShapeType', '')}への変更は未実施（5-3）",
                "    - 箇条書きはプレースホルダに配置",
            ]

        elif no == 5:
            lines += [
                f"    - 挿入→図形「{task_params.get('task5_2', {}).get('rectShapeInsertName', '')}」3個",
                "    - 大2個は同じ幅、小1個は明らかに幅が狭い（5-2未実施）",
                "    - 3個を横並びまたは縦並びで視認できるよう配置",
                "    - 箇条書きはプレースホルダに配置",
            ]

        elif no == 6:
            labels = p4.get("labels", [])
            gate = task_params.get("task5_5", {})
            if labels:
                lines.append(
                    f"    - 挿入→図形でラベル付き3個「{'」「'.join(labels)}」を重ねて配置"
                )
                lines.append(
                    f"      重なり順は未調整（5-4未実施）。手前からの順序は意図的にバラバラ"
                )
            lines += [
                f"    - 挿入→図形「{gate.get('gateShapeInsertName', '')}」3個を別グループとして近接配置",
                "    - 3個のゲート図形はグループ化未実施（5-5）",
                "    - 箇条書きはプレースホルダに配置",
            ]

        else:
            lines.append("    - 箇条書きはタイトル下プレースホルダに配置")

        lines.append("")

    return lines


def _layout_p6(slide_map: list[dict], task_params: dict) -> list[str]:
    custom = task_params.get("task6_4", {}).get("customShowName", "")
    target_slides = task_params.get("task6_4", {}).get("slideNos", [4, 5, 6])
    lines = _header()
    lines += [
        "  全スライド共通:",
        "    - レイアウト「タイトルとコンテンツ」。箇条書きはプレースホルダに転記",
        "    - スピーカーノートは各スライドのノート欄に貼り付け用データどおり入力（6-6印刷用）",
        "    - 図形・画像の追加は不要（操作はファイル設定・印刷が中心）",
        "",
    ]
    for s in slide_map:
        no = s["slideNo"]
        extra = ""
        if no in target_slides:
            extra = f" — 目的別スライドショー「{custom}」の対象（6-4未作成）"
        lines.append(f"  スライド{no}「{s['title']}」{extra}")
        lines.append("    - 箇条書き＋スピーカーノートを配置")
        lines.append("")

    return lines


def _layout_p7(slide_map: list[dict], task_params: dict) -> list[str]:
    p1 = task_params.get("task7_1", {})
    p2 = task_params.get("task7_2", {})
    lines = _header()
    lines += [
        "  全スライド共通:",
        "    - 箇条書きはタイトル下プレースホルダに転記",
        "    - コメント・ハイパーリンク・フッター・アウトライン挿入は未実施",
        "",
    ]
    for s in slide_map:
        no = s["slideNo"]
        title = s["title"]
        lines.append(f"  スライド{no}「{title}」")
        if no == 1:
            phrase = p2.get("hyperlinkPhrase", "")
            lines += [
                f"    - タイトル下またはサブタイトル付近に文字列「{phrase}」を配置（ハイパーリンク未設定・7-2未実施）",
                f"    - コメント「{p1.get('commentText', '')}」は未投稿（7-1未実施）",
                "    - 箇条書きはプレースホルダに配置",
            ]
        elif no in (5, 6):
            footer = task_params.get("task7_5", {}).get("footerText", "")
            lines += [
                f"    - 事例スライド。フッター「{footer}」は未設定（7-5未実施）",
                "    - 箇条書きはプレースホルダに配置",
            ]
        elif no == 8:
            doc = task_params.get("task7_3", {}).get("outlineDoc", "")
            lines += [
                f"    - アウトライン「{doc}」からの挿入スライドは未挿入（7-3未実施）",
                "    - 現時点では通常コンテンツスライドとして箇条書きのみ配置",
            ]
        else:
            lines.append("    - グローバルフッター・スライド番号は未設定（7-4未実施）")
            lines.append("    - 箇条書きはプレースホルダに配置")
        lines.append("")
    return lines


def _layout_p8(slide_map: list[dict], task_params: dict) -> list[str]:
    p1 = task_params.get("task8_1", {})
    p2 = task_params.get("task8_2", {})
    tbl_title = p1.get("slideTitle", "")
    lines = _header()
    lines += [
        "  全スライド共通:",
        "    - 箇条書きはプレースホルダに転記",
        "    - 表スタイル・背景色・スライドサイズ・グレースケールは未設定",
        "",
    ]
    for s in slide_map:
        no = s["slideNo"]
        title = s["title"]
        lines.append(f"  スライド{no}「{title}」")
        if no == 1:
            lines += [
                f"    - タイトルスライド。背景色「{p2.get('backgroundColor', '')}」は未設定（8-2未実施）",
                "    - 箇条書きまたはサブタイトルを配置",
            ]
        elif title == tbl_title:
            lines += [
                "    - 挿入→表（行4列3程度）と挿入→イラストを同一スライドに配置",
                f"    - 表スタイル「{p1.get('tableStyle', '')}」は未変更（8-1未実施）",
                "    - イラストのグレースケールは未設定（8-5未実施）",
                "    - 箇条書きは表・イラストと重ならない位置のプレースホルダに配置",
            ]
        else:
            lines.append("    - 箇条書きはプレースホルダに配置")
        lines.append("")
    return lines


def _layout_p9(slide_map: list[dict], task_params: dict) -> list[str]:
    p1 = task_params.get("task9_1", {})
    p2 = task_params.get("task9_2", {})
    p3 = task_params.get("task9_3", {})
    p4 = task_params.get("task9_4", {})
    video_title = p1.get("slideTitle", "")
    video_file = p1.get("videoFile", "")
    lines = _header()
    lines += [
        "  ※ メディアは「挿入→メディアのこのデバイス」前提。Copilot生成イラストで代用禁止。",
        "  ※ 第1段階では本文・表データまで作成し、メディア挿入は手動後工程でも可。",
        "",
    ]
    for s in slide_map:
        no = s["slideNo"]
        title = s["title"]
        lines.append(f"  スライド{no}「{title}」")
        if no == 1:
            lines += [
                "    - 挿入→オーディオのアイコンをスライドに配置（または未挿入のままアイコン位置を空けない）",
                f"    - スライド切替後再生・フェード{p3.get('fadeOutSec', '')}秒は未設定（9-3未実施）",
                "    - 箇条書きはプレースホルダに配置",
            ]
        elif no == 2:
            rows = p4.get("tableRows", [])
            cat = p4.get("categoryColumn", "")
            dat = p4.get("dataColumn", "")
            row_desc = " / ".join(f"{r[0]}={r[1]}" for r in rows) if rows else ""
            lines += [
                f"    - 挿入→表: 列「{cat}」「{dat}」、行データ例: {row_desc}",
                "    - 集合縦棒グラフは未作成（9-4未実施）。凡例ありの初期状態でも可",
                "    - データテーブル・凡例削除は未実施（9-5未実施）",
                "    - 箇条書きは表と別プレースホルダに配置",
            ]
        elif title == video_title:
            lines += [
                f"    - ビデオ「{video_file}」は未挿入または未トリミング（9-1/9-2未実施）",
                f"    - トリム開始{p2.get('trimStartSec', '')}秒・終了{p2.get('trimEndSec', '')}秒は未設定",
                "    - 挿入時はアイコン表示を使用",
                "    - 箇条書きはプレースホルダに配置",
            ]
        else:
            lines.append("    - 箇条書きはプレースホルダに配置")
        lines.append("")
    return lines


def _layout_p10(slide_map: list[dict], task_params: dict) -> list[str]:
    p1 = task_params.get("task10_1", {})
    p6 = task_params.get("task10_6", {})
    p8 = task_params.get("task10_8", {})
    lines = _header()
    lines += [
        "  ※ スライドマスター・配布資料マスターの操作は受験者タスク。初期状態はマスター未変更。",
        "  ※ 第1段階: 通常スライド本文のみ作成。マスター操作は手動後工程。",
        "",
        "  マスター未変更の状態（全タスク未実施）:",
        f"    - テーマ「{p1.get('masterTheme', '')}」未適用（10-1）",
        "    - スライド番号・フッター・カスタムレイアウトはデフォルトのまま",
        f"    - 配布資料フッター「{p8.get('handoutFooter', '')}」未設定（10-8）",
        f"    - カスタムレイアウト「{p6.get('customLayoutName', '')}」未作成（10-6）",
        "",
    ]
    for s in slide_map:
        no = s["slideNo"]
        title = s["title"]
        layout = s.get("layout", "タイトルとコンテンツ")
        extra = ""
        if no == 2:
            extra = " — レイアウト「2 つのコンテンツ」（10-4背景非表示は未実施）"
        lines.append(f"  スライド{no}「{title}」{extra}")
        lines.append(f"    - レイアウト: {layout}")
        lines.append("    - 箇条書きはプレースホルダに転記（図形・画像の追加は不要）")
        lines.append("")
    return lines
