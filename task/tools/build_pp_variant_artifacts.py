"""Generate MD / app JSON from PowerPoint variant JSON. Validate consistency."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
TOOLS = Path(__file__).resolve().parent
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
from ppt_slide_outline import render_slide_map_md
from ppt_copilot_content import (
    content_md_path,
    expected_slide_count,
    parse_content_md,
    slide_titles_from_json_sets,
    validate_content_md,
)

FIXED_CONTENT_PROJECTS = frozenset({4, 5, 6, 7, 8, 9, 10})

PROJECT_CONFIG = {
    1: {"task_count": 8, "forbidden": set()},
    2: {"task_count": 8, "forbidden": set()},
    3: {"task_count": 7, "forbidden": set()},
    4: {"task_count": 8, "forbidden": set()},
    5: {"task_count": 7, "forbidden": set()},
    6: {"task_count": 7, "forbidden": set()},
    7: {"task_count": 5, "forbidden": set()},
    8: {"task_count": 5, "forbidden": set()},
    9: {"task_count": 5, "forbidden": set()},
    10: {"task_count": 8, "forbidden": set()},
}

PROJECT_OVERVIEW_OPS = {
    1: [
        ("1-1", "レイアウト付きスライド挿入"),
        ("1-2", "スライド非表示"),
        ("1-3", "レイアウト変更＋プレースホルダ入力"),
        ("1-4", "箇条書き2段組み"),
        ("1-5", "文字間隔（幅pt）"),
        ("1-6", "セクション追加"),
        ("1-7", "セクション名変更"),
        ("1-8", "サマリーズーム"),
    ],
    2: [
        ("2-1", "全スライド画面切り替え"),
        ("2-2", "画面切り替え継続時間"),
        ("2-3", "個別スライド画面切り替え"),
        ("2-4", "自動タイミング"),
        ("2-5", "3Dモデルアニメーション"),
        ("2-6", "画像登場アニメーション"),
        ("2-7", "箇条書きアニメ効果変更"),
        ("2-8", "アニメーション軌跡"),
    ],
    3: [
        ("3-1", "SmartArt手順追加"),
        ("3-2", "SmartArt色変更"),
        ("3-3", "箇条書き→SmartArt変換"),
        ("3-4", "3Dモデル挿入"),
        ("3-5", "3Dモデルビュー・サイズ"),
        ("3-6", "スライドズーム"),
        ("3-7", "セクションズーム"),
    ],
    4: [
        ("4-1", "図形への文字入力"),
        ("4-2", "文字の塗りつぶし"),
        ("4-3", "図の効果・光彩"),
        ("4-4", "クイックスタイル＋アート効果"),
        ("4-5", "画像上揃え"),
        ("4-6", "画像トリミング"),
        ("4-7", "図形塗り＋枠線"),
        ("4-8", "テキストボックス上下中央"),
    ],
    5: [
        ("5-1", "図形右揃え"),
        ("5-2", "図形幅の統一"),
        ("5-3", "図形の変更"),
        ("5-4", "重なり順変更"),
        ("5-5", "図形グループ化"),
        ("5-6", "代替テキスト装飾化"),
        ("5-7", "アイコン塗りつぶし"),
    ],
    6: [
        ("6-1", "ドキュメント検査"),
        ("6-2", "読み取り専用"),
        ("6-3", "自動プレゼンテーション"),
        ("6-4", "目的別スライドショー"),
        ("6-5", "アウトライン印刷"),
        ("6-6", "ノート印刷"),
        ("6-7", "配布資料印刷設定"),
    ],
    7: [
        ("7-1", "コメント挿入"),
        ("7-2", "ハイパーリンク設定"),
        ("7-3", "アウトライン挿入"),
        ("7-4", "フッター（スライド番号＋ドメイン）"),
        ("7-5", "事例スライド専用フッター"),
    ],
    8: [
        ("8-1", "表スタイル変更"),
        ("8-2", "スライド背景色"),
        ("8-3", "スライドサイズ16:9"),
        ("8-4", "スライドサイズカスタム"),
        ("8-5", "グレースケール表示・イラスト"),
    ],
    9: [
        ("9-1", "ビデオ挿入"),
        ("9-2", "ビデオトリミング"),
        ("9-3", "オーディオ再生設定"),
        ("9-4", "集合縦棒グラフ作成"),
        ("9-5", "グラフデータテーブル"),
    ],
    10: [
        ("10-1", "スライドマスターテーマ"),
        ("10-2", "マスターにスライド番号"),
        ("10-3", "タイトルスライド番号非表示"),
        ("10-4", "2つのコンテンツ背景非表示"),
        ("10-5", "マスターフッター削除"),
        ("10-6", "カスタムレイアウト作成"),
        ("10-7", "配布資料マスター日付削除"),
        ("10-8", "配布資料マスターフッター"),
    ],
}


def _check_unique(values: list, label: str, issues: list[str]) -> None:
    if len(values) != len(set(str(v) for v in values)):
        issues.append(f"{label} reused across sets: {values}")


def _collect_text_blob(s: dict) -> str:
    parts = [s.get("problemStatement", "")]
    parts.extend(s.get("tasks", []))
    for slide in s.get("contentBlocks", {}).get("slideMap", []):
        parts.append(slide.get("title", ""))
        parts.extend(slide.get("bullets", []) or [])
    return "\n".join(parts)


def validate_project_1(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    defaults = data["variantRules"].get("textbookDefaults", {})
    for s in sets:
        p5 = s["taskParams"]["task1_5"]
        if p5.get("spacingPt") == defaults.get("spacingPt", 4):
            issues.append(f"set{s['setNo']}: spacingPt must differ from textbook 4")
        p1 = s["taskParams"]["task1_1"]
        if p1.get("layoutUi") == defaults.get("layoutInsert", "テーブルスライド"):
            issues.append(f"set{s['setNo']}: layoutInsert must differ from textbook")
        for title in s["taskParams"]["task1_8"]["linkTitles"]:
            if not _title_in_map(s, title):
                issues.append(f"set{s['setNo']}: zoom link title missing in slideMap: {title}")
    _check_unique([s["taskParams"]["task1_5"]["spacingPt"] for s in sets], "spacingPt", issues)
    _check_unique([s["taskParams"]["task1_8"]["zoomTitle"] for s in sets], "zoomTitle", issues)
    _check_unique([s["taskParams"]["task1_6"]["sectionName"] for s in sets], "sectionName", issues)
    return issues


def validate_project_2(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    defaults = data["variantRules"].get("textbookDefaults", {})
    for s in sets:
        p2 = s["taskParams"]["task2_2"]
        p4 = s["taskParams"]["task2_4"]
        p1 = s["taskParams"]["task2_1"]
        if p2.get("durationSec") == defaults.get("durationSec", 3):
            issues.append(f"set{s['setNo']}: durationSec must differ from textbook 3")
        if p4.get("autoAdvanceSec") == defaults.get("autoAdvanceSec", 5):
            issues.append(f"set{s['setNo']}: autoAdvanceSec must differ from textbook 5")
        if p1.get("transition") == defaults.get("transitionAll", "スプリット"):
            issues.append(f"set{s['setNo']}: transitionAll must differ from textbook スプリット")
        title = s["taskParams"]["task2_7"]["slideTitle"]
        if not _title_in_map(s, title):
            issues.append(f"set{s['setNo']}: slide title missing: {title}")
    _check_unique([s["taskParams"]["task2_2"]["durationSec"] for s in sets], "durationSec", issues)
    _check_unique([s["taskParams"]["task2_1"]["transition"] for s in sets], "transitionAll", issues)
    return issues


def validate_project_3(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    defaults = data["variantRules"].get("textbookDefaults", {})
    for s in sets:
        p4 = s["taskParams"]["task3_4"]
        p5 = s["taskParams"]["task3_5"]
        p1 = s["taskParams"]["task3_1"]
        if p4.get("width") == defaults.get("modelWidth", 2.5):
            issues.append(f"set{s['setNo']}: model width must differ from textbook 2.5")
        if p5.get("height") == defaults.get("modelHeight", 6.5):
            issues.append(f"set{s['setNo']}: model height must differ from textbook 6.5")
        if p1.get("smartArtType") == defaults.get("smartArtType", "タイムライン"):
            issues.append(f"set{s['setNo']}: smartArtType must differ from タイムライン")
        for t in s["taskParams"]["task3_6"]["linkTitles"]:
            if not _title_in_map(s, t):
                issues.append(f"set{s['setNo']}: slide zoom target missing: {t}")
    _check_unique([s["taskParams"]["task3_2"]["smartArtColor"] for s in sets], "smartArtColor", issues)
    return issues


def _title_in_map(s: dict, title: str) -> bool:
    for slide in s.get("contentBlocks", {}).get("slideMap", []):
        if slide.get("title") == title:
            return True
    return False


def _shape_by_id(s: dict) -> dict[str, dict]:
    out: dict[str, dict] = {}
    for slide in s.get("contentBlocks", {}).get("slideMap", []):
        for sh in slide.get("shapes", []) or []:
            sid = sh.get("id")
            if sid:
                out[sid] = sh
    return out


def _shape_consistency_p4(s: dict, issues: list[str]) -> None:
    tp = s["taskParams"]
    shapes = _shape_by_id(s)
    checks = [
        ("inputShape", tp["task4_1"]["inputShapeType"], "task4_1"),
        ("textBoxShape", tp["task4_7"]["textBoxShapeType"], "task4_7"),
        ("centerTextBox", tp["task4_7"]["textBoxShapeType"], "task4_8"),
    ]
    for shape_id, expected_type, task_key in checks:
        sh = shapes.get(shape_id)
        if not sh:
            issues.append(f"set{s['setNo']}: slideMap missing shape id={shape_id} for {task_key}")
        elif sh.get("type") != expected_type:
            issues.append(
                f"set{s['setNo']}: {shape_id} type {sh.get('type')!r} != taskParams {expected_type!r}"
            )
    blob = "\n".join(s.get("tasks", []))
    for key in ("inputShapeType", "textBoxShapeType"):
        val = tp.get("task4_1" if key == "inputShapeType" else "task4_7", {}).get(key)
        if val and val not in blob:
            issues.append(f"set{s['setNo']}: tasks missing shape type {val!r} ({key})")


def _shape_consistency_p5(s: dict, issues: list[str]) -> None:
    tp = s["taskParams"]
    shapes = _shape_by_id(s)
    if shapes.get("sourceShape", {}).get("type") != tp["task5_3"]["sourceShapeType"]:
        issues.append(f"set{s['setNo']}: sourceShape type mismatch task5_3")
    for n in range(1, 5):
        sid = f"alignShape{n}"
        if shapes.get(sid, {}).get("type") != tp["task5_1"]["alignShapeType"]:
            issues.append(f"set{s['setNo']}: {sid} type mismatch task5_1")
    for sid in ("rectLarge1", "rectLarge2", "rectSmall"):
        if shapes.get(sid, {}).get("type") != tp["task5_2"]["rectShapeType"]:
            issues.append(f"set{s['setNo']}: {sid} type mismatch task5_2")
    for n in range(1, 4):
        sid = f"gateShape{n}"
        if shapes.get(sid, {}).get("type") != tp["task5_5"]["gateShapeType"]:
            issues.append(f"set{s['setNo']}: {sid} type mismatch task5_5")
    blob = "\n".join(s.get("tasks", []))
    for val in (
        tp["task5_1"]["alignShapeType"],
        tp["task5_2"]["rectShapeType"],
        tp["task5_3"]["sourceShapeType"],
        tp["task5_3"]["targetShapeType"],
        tp["task5_5"]["gateShapeType"],
    ):
        if val not in blob:
            issues.append(f"set{s['setNo']}: tasks missing shape type {val!r}")
    title = tp["task5_6"]["slideTitle"]
    if not _title_in_map(s, title):
        issues.append(f"set{s['setNo']}: decorative slide title missing in slideMap: {title}")


def validate_project_4(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    defaults = data["variantRules"].get("textbookDefaults", {})
    for s in sets:
        tp = s["taskParams"]
        checks = [
            (tp["task4_1"]["inputShapeType"], defaults.get("inputShapeType"), "inputShapeType"),
            (tp["task4_1"]["shapeText"], defaults.get("shapeText"), "shapeText"),
            (tp["task4_2"]["targetPhrase"], defaults.get("targetPhrase"), "targetPhrase"),
            (tp["task4_2"]["fillColor"], defaults.get("fillColor"), "fillColor"),
            (tp["task4_3"]["glowPt"], defaults.get("glowPt"), "glowPt"),
            (tp["task4_4"]["quickStyle"], defaults.get("quickStyle"), "quickStyle"),
            (tp["task4_4"]["artEffect"], defaults.get("artEffect"), "artEffect"),
            (tp["task4_7"]["borderPt"], defaults.get("borderPt"), "borderPt"),
        ]
        for actual, default, label in checks:
            if default is not None and actual == default:
                issues.append(f"set{s['setNo']}: {label} must differ from textbook {default!r}")
        _shape_consistency_p4(s, issues)
    _check_unique([s["taskParams"]["task4_1"]["inputShapeType"] for s in sets], "inputShapeType", issues)
    _check_unique([s["taskParams"]["task4_7"]["textBoxShapeType"] for s in sets], "textBoxShapeType", issues)
    _check_unique([s["taskParams"]["task4_1"]["shapeText"] for s in sets], "shapeText", issues)
    _check_unique([s["taskParams"]["task4_2"]["targetPhrase"] for s in sets], "targetPhrase", issues)
    _check_unique([s["taskParams"]["task4_2"]["fillColor"] for s in sets], "fillColor", issues)
    _check_unique([s["taskParams"]["task4_3"]["glowPt"] for s in sets], "glowPt", issues)
    _check_unique([s["taskParams"]["task4_4"]["quickStyle"] for s in sets], "quickStyle", issues)
    _check_unique([s["taskParams"]["task4_4"]["artEffect"] for s in sets], "artEffect", issues)
    _check_unique([s["taskParams"]["task4_7"]["borderPt"] for s in sets], "borderPt", issues)
    return issues


def validate_project_5(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    defaults = data["variantRules"].get("textbookDefaults", {})
    for s in sets:
        tp = s["taskParams"]
        if tp["task5_3"]["targetShapeType"] == defaults.get("targetShapeType"):
            issues.append(f"set{s['setNo']}: targetShapeType must differ from textbook スマイル")
        if tp["task5_1"]["alignShapeType"] == defaults.get("alignShapeType"):
            issues.append(f"set{s['setNo']}: alignShapeType must differ from textbook 丸")
        if tp["task5_7"]["iconFillColor"] == defaults.get("iconFillColor"):
            issues.append(f"set{s['setNo']}: iconFillColor must differ from textbook 濃い赤")
        if tp["task5_3"]["sourceShapeType"] == defaults.get("sourceShapeType"):
            issues.append(f"set{s['setNo']}: sourceShapeType must differ from textbook 星")
        _shape_consistency_p5(s, issues)
    _check_unique([s["taskParams"]["task5_1"]["alignShapeType"] for s in sets], "alignShapeType", issues)
    _check_unique([s["taskParams"]["task5_2"]["rectShapeType"] for s in sets], "rectShapeType", issues)
    _check_unique([s["taskParams"]["task5_3"]["sourceShapeType"] for s in sets], "sourceShapeType", issues)
    _check_unique([s["taskParams"]["task5_3"]["targetShapeType"] for s in sets], "targetShapeType", issues)
    _check_unique(
        ["|".join(s["taskParams"]["task5_4"]["labels"]) for s in sets], "zOrderLabels", issues
    )
    _check_unique([s["taskParams"]["task5_6"]["slideTitle"] for s in sets], "decorativeSlideTitle", issues)
    _check_unique([s["taskParams"]["task5_7"]["iconFillColor"] for s in sets], "iconFillColor", issues)
    return issues


def validate_project_6(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    defaults = data["variantRules"].get("textbookDefaults", {})
    for s in sets:
        tp = s["taskParams"]
        if tp["task6_4"]["customShowName"] == defaults.get("customShowName"):
            issues.append(f"set{s['setNo']}: customShowName must differ from textbook")
        if tp["task6_5"]["outlineCopies"] == defaults.get("outlineCopies"):
            issues.append(f"set{s['setNo']}: outlineCopies must differ from textbook 6")
        if tp["task6_6"]["notesCopies"] == defaults.get("notesCopies"):
            issues.append(f"set{s['setNo']}: notesCopies must differ from textbook 3")
        if tp["task6_7"]["handoutCopies"] == defaults.get("handoutCopies"):
            issues.append(f"set{s['setNo']}: handoutCopies must differ from textbook 4")
    _check_unique([s["taskParams"]["task6_4"]["customShowName"] for s in sets], "customShowName", issues)
    _check_unique([s["taskParams"]["task6_5"]["outlineCopies"] for s in sets], "outlineCopies", issues)
    _check_unique([s["taskParams"]["task6_6"]["notesCopies"] for s in sets], "notesCopies", issues)
    _check_unique([s["taskParams"]["task6_7"]["handoutCopies"] for s in sets], "handoutCopies", issues)
    return issues


def validate_project_7(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    defaults = data["variantRules"].get("textbookDefaults", {})
    for s in sets:
        tp = s["taskParams"]
        if tp["task7_2"]["hyperlinkUrl"] == defaults.get("hyperlinkUrl"):
            issues.append(f"set{s['setNo']}: hyperlinkUrl must differ from textbook")
        if tp["task7_4"]["footerDomain"] == defaults.get("footerDomain"):
            issues.append(f"set{s['setNo']}: footerDomain must differ from textbook")
        if tp["task7_1"]["commentText"] == defaults.get("commentText"):
            issues.append(f"set{s['setNo']}: commentText must differ from textbook")
        for title in tp["task7_5"]["slideTitles"]:
            if not _title_in_map(s, title):
                issues.append(f"set{s['setNo']}: case slide title missing in slideMap: {title}")
    _check_unique([s["taskParams"]["task7_2"]["hyperlinkUrl"] for s in sets], "hyperlinkUrl", issues)
    _check_unique([s["taskParams"]["task7_4"]["footerDomain"] for s in sets], "footerDomain", issues)
    _check_unique([s["taskParams"]["task7_5"]["footerText"] for s in sets], "caseFooter", issues)
    _check_unique(
        ["|".join(s["taskParams"]["task7_5"]["slideTitles"]) for s in sets],
        "caseSlideTitles",
        issues,
    )
    return issues


def validate_project_8(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    defaults = data["variantRules"].get("textbookDefaults", {})
    for s in sets:
        tp = s["taskParams"]
        if tp["task8_1"]["tableStyle"] == defaults.get("tableStyle"):
            issues.append(f"set{s['setNo']}: tableStyle must differ from textbook")
        if tp["task8_2"]["backgroundColor"] == defaults.get("backgroundColor"):
            issues.append(f"set{s['setNo']}: backgroundColor must differ from textbook")
        if tp["task8_4"]["widthCm"] == defaults.get("slideWidthCm"):
            issues.append(f"set{s['setNo']}: slideWidthCm must differ from textbook")
        if tp["task8_4"]["heightCm"] == defaults.get("slideHeightCm"):
            issues.append(f"set{s['setNo']}: slideHeightCm must differ from textbook")
        title = tp["task8_1"]["slideTitle"]
        if not _title_in_map(s, title):
            issues.append(f"set{s['setNo']}: table slide title missing: {title}")
    _check_unique([s["taskParams"]["task8_1"]["tableStyle"] for s in sets], "tableStyle", issues)
    _check_unique([s["taskParams"]["task8_2"]["backgroundColor"] for s in sets], "backgroundColor", issues)
    _check_unique(
        [f"{s['taskParams']['task8_4']['widthCm']}x{s['taskParams']['task8_4']['heightCm']}" for s in sets],
        "slideSize",
        issues,
    )
    return issues


def validate_project_9(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    defaults = data["variantRules"].get("textbookDefaults", {})
    for s in sets:
        tp = s["taskParams"]
        if tp["task9_2"]["trimStartSec"] == defaults.get("trimStartSec"):
            issues.append(f"set{s['setNo']}: trimStartSec must differ from textbook")
        if tp["task9_2"]["trimEndSec"] == defaults.get("trimEndSec"):
            issues.append(f"set{s['setNo']}: trimEndSec must differ from textbook")
        if tp["task9_4"]["categoryColumn"] == defaults.get("categoryColumn"):
            issues.append(f"set{s['setNo']}: categoryColumn must differ from textbook")
        vt = tp["task9_1"]["slideTitle"]
        if not _title_in_map(s, vt):
            issues.append(f"set{s['setNo']}: video slide title missing: {vt}")
    _check_unique(
        [f"{s['taskParams']['task9_2']['trimStartSec']}-{s['taskParams']['task9_2']['trimEndSec']}" for s in sets],
        "videoTrim",
        issues,
    )
    _check_unique([s["taskParams"]["task9_4"]["categoryColumn"] for s in sets], "chartCategoryColumn", issues)
    _check_unique([s["taskParams"]["task9_1"]["videoFile"] for s in sets], "videoFile", issues)
    return issues


def validate_project_10(data: dict) -> list[str]:
    issues: list[str] = []
    sets = data["sets"]
    defaults = data["variantRules"].get("textbookDefaults", {})
    for s in sets:
        tp = s["taskParams"]
        if tp["task10_1"]["masterTheme"] == defaults.get("masterTheme"):
            issues.append(f"set{s['setNo']}: masterTheme must differ from textbook")
        if tp["task10_8"]["handoutFooter"] == defaults.get("handoutFooter"):
            issues.append(f"set{s['setNo']}: handoutFooter must differ from textbook")
        if tp["task10_6"]["customLayoutName"] == defaults.get("customLayoutName"):
            issues.append(f"set{s['setNo']}: customLayoutName must differ from textbook")
    _check_unique([s["taskParams"]["task10_1"]["masterTheme"] for s in sets], "masterTheme", issues)
    _check_unique([s["taskParams"]["task10_8"]["handoutFooter"] for s in sets], "handoutFooterText", issues)
    _check_unique([s["taskParams"]["task10_6"]["customLayoutName"] for s in sets], "customLayoutName", issues)
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


def validate_common(data: dict, cfg: dict) -> list[str]:
    issues: list[str] = []
    forbidden = set(data.get("variantRules", {}).get("forbiddenTexts", [])) | cfg.get("forbidden", set())
    task_count = cfg["task_count"]
    for s in data["sets"]:
        if len(s["tasks"]) != task_count:
            issues.append(f"set{s['setNo']}: task count {len(s['tasks'])} != {task_count}")
        blob = _collect_text_blob(s)
        for f in forbidden:
            if f and f in blob:
                issues.append(f"set{s['setNo']}: forbidden text found: {f}")
        sm = s.get("contentBlocks", {}).get("slideMap", [])
        if not sm:
            issues.append(f"set{s['setNo']}: slideMap is empty")
    return issues


def validate_copilot_content(data: dict) -> list[str]:
    pid = data["projectId"]
    if pid not in FIXED_CONTENT_PROJECTS:
        return []
    path = content_md_path(pid)
    if not path.exists():
        return [f"content: missing {path.name}"]
    content = parse_content_md(path)
    forbidden = set(data.get("variantRules", {}).get("forbiddenTexts", []))
    return validate_content_md(
        pid,
        content,
        expected_sets=len(data["sets"]),
        expected_slides_per_set=expected_slide_count(data["sets"]),
        slide_titles_by_set=slide_titles_from_json_sets(data["sets"]),
        forbidden_texts=forbidden,
        require_speaker_notes=(pid == 6),
        min_bullets=3,
    )


def validate(data: dict, cfg: dict) -> list[str]:
    issues = validate_common(data, cfg)
    pid = data["projectId"]
    validator = PROJECT_VALIDATORS.get(pid)
    if validator:
        issues.extend(validator(data))
    issues.extend(validate_copilot_content(data))
    return issues


def write_app_json(data: dict) -> dict:
    pid = data["projectId"]
    app_sets = []
    for s in data["sets"]:
        items = []
        for i, t in enumerate(s["tasks"], 1):
            desc = re.sub(rf"^タスク{pid}-\d+　", "", t)
            if i == 1 and "\n" in desc:
                pass
            elif i == 1:
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


def write_md_rich(data: dict) -> str:
    pid = data["projectId"]
    forbidden = data.get("variantRules", {}).get("forbiddenTexts", [])
    lines = [
        f"# MOSスペシャリスト PowerPoint 類題 — Project{pid} 配置別（5セット）",
        "",
        f"元ファイル {data['sourceFile']} の操作をもとに、",
        "レイアウトを差別化した 5 セットの類題プレゼン用問題文です。",
        "",
        "**関連ファイル:**",
        f"- 正本 JSON: `task/類題Json/MOS_PowerPoint類題_project{pid}_配置別_5セット_問題文.json`",
        f"- アプリ用 JSON: `task/類題Json/MOS_PowerPoint類題_project{pid}_配置別_5セット_問題文_アプリ用.json`",
        f"- Copilot 用: `task/PowerPoint_類題_Copilot_project{pid}_配置別.md`",
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
    lines += [
        "",
        f"**禁止文字列（教材流用不可）:** {' / '.join(forbidden[:8])}{' ...' if len(forbidden) > 8 else ''}",
        "",
        "### Copilot で pptx を作るときの注意",
        "",
        "1. **slideMap の全スライドを事前配置**し、タスク対象オブジェクトを含める",
        "2. **受験者が行う操作は事前に完了させない**（preTaskState 参照）",
        "3. **空プレースホルダ・ダミー文字を入れない**",
        "4. 詳細は `task/PowerPoint_類題_Copilot_project{pid}_配置別.md` を使用",
        "",
        "---",
        "",
    ]
    for s in data["sets"]:
        ly = s["layout"]
        lines += [
            f"## セット{s['setNo']}：{s['workbook']}",
            "",
            f"**テーマ:** {s['theme']}  ",
            f"**レイアウト:** サイズ={ly.get('slideSize')} / デザイン={ly.get('designTheme')} / "
            f"バリアント={ly.get('variant')} / 色={ly.get('colorTheme')} / フォント={ly.get('fontTheme')}  ",
            f"**スライド枚数:** {ly.get('slideCount')}  ",
            "",
            "### 問題文",
            "",
            f"{s['problemStatement']}  ",
            "問題）  ",
        ]
        for t in s["tasks"]:
            lines.append(f"{t}  ")
        lines += [
            "",
            "### スライド構成（事前配置データ）",
            "",
            "以下を **すべて** プレゼンに配置してください。",
            "",
        ]
        cb = s["contentBlocks"]
        lines.extend(render_slide_map_md(cb.get("slideMap", []), cb.get("preTaskState")))
        pre = cb.get("preTaskState", {})
        lines += ["", "**受験者操作前の状態:**"]
        for k, v in pre.items():
            if v in (False, [], None, ""):
                lines.append(f"- {k}: 未実施")
            elif v is True:
                lines.append(f"- {k}: 済（要確認）")
        lines += ["", "---", ""]
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project", type=int, choices=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10], required=True)
    parser.add_argument("--validate-only", action="store_true")
    args = parser.parse_args()
    pid = args.project
    json_path = BASE / "類題Json" / f"MOS_PowerPoint類題_project{pid}_配置別_5セット_問題文.json"
    if not json_path.exists():
        print(f"MISSING {json_path}")
        sys.exit(1)
    data = json.loads(json_path.read_text(encoding="utf-8"))
    cfg = PROJECT_CONFIG[pid]
    cfg = {**cfg, "forbidden": set(data.get("variantRules", {}).get("forbiddenTexts", []))}
    issues = validate(data, cfg)
    label = f"PP Project{pid}"
    if issues:
        print(f"{label}: FAIL {issues}")
        sys.exit(1)
    print(f"{label}: OK")
    if args.validate_only:
        return
    md_path = BASE / f"MOS_PowerPoint類題_project{pid}_配置別_5セット_問題文.md"
    app_path = BASE / "類題Json" / f"MOS_PowerPoint類題_project{pid}_配置別_5セット_問題文_アプリ用.json"
    md_path.write_text(write_md_rich(data), encoding="utf-8")
    app_path.write_text(
        json.dumps(write_app_json(data), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"wrote {md_path.name}")
    print(f"wrote {app_path.name}")


if __name__ == "__main__":
    main()
