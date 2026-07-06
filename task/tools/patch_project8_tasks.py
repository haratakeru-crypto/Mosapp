"""Patch Project8 Word variant JSON: diversify tasks 8-1..8-7 away from textbook patterns."""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
JSON_PATH = BASE / "類題Json" / "MOS_Word類題_project8_配置別_5セット_問題文.json"
TOOLS = Path(__file__).resolve().parent
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
from word_document_outline import apply_outlines_to_json

SET_PATCHES = [
    {
        "setNo": 1,
        "fn_heading": "POSセキュリティの基本",
        "bib_heading": "資料一覧",
        "task8_1": {
            "subtitle": "－店舗POSの守り方－",
            "tocStyle": "自動作成の目次2",
            "tocPlacement": "afterSubtitle",
            "insertInShape": False,
            "placementDescription": "副題の直後（図形外）",
        },
        "task8_2": {"heading": "POSセキュリティの基本", "footnoteAnchor": "レジ端末", "footnoteText": "不正操作を検知する仕組み"},
        "task8_3": {"footnoteStartNumber": "*", "numberFormat": "Symbol", "numberLabel": "*（アスタリスク）"},
        "task8_4": {"heading": "資料一覧", "specialChar": "©", "specialCharName": "©（著作権記号）"},
        "task8_5": {"page": "last", "smartArtType": "基本プロセス", "smartArtLabels": ["注文", "在庫", "決済"]},
        "task8_6": {"smartArtColor": "モノクロ", "smartArtAccent": ""},
        "task8_7": {
            "linkShapeText": "一覧へ",
            "linkType": "heading",
            "linkTargetHeading": "資料一覧",
            "linkShapePlacement": "underHeading",
            "linkNearHeading": "対策の基本",
        },
        "link_shape_note": "見出し「対策の基本」下に「一覧へ」図形（リンク未設定）",
    },
    {
        "setNo": 2,
        "fn_heading": "カルテ保護の要点",
        "bib_heading": "文献リスト",
        "task8_1": {
            "subtitle": "－電子カルテの守り方－",
            "tocStyle": "自動作成の目次2",
            "tocPlacement": "underHeading",
            "tocAnchorHeading": "はじめに",
            "insertInShape": False,
            "placementDescription": "見出し「はじめに」の下",
        },
        "task8_2": {"heading": "カルテ保護の要点", "footnoteAnchor": "電子カルテ", "footnoteText": "患者情報を保護する規格"},
        "task8_3": {"footnoteStartNumber": "壱", "numberFormat": "Chosho", "numberLabel": "壱"},
        "task8_4": {"heading": "文献リスト", "specialChar": "™", "specialCharName": "™（商標記号）"},
        "task8_5": {"page": "last", "smartArtType": "サイクル", "smartArtLabels": ["患者", "診療", "連携"]},
        "task8_6": {"smartArtColor": "アクセント1", "smartArtAccent": ""},
        "task8_7": {
            "linkShapeText": "本院案内",
            "linkType": "url",
            "linkUrl": "https://medical-hospital.example.jp/security",
            "linkShapePlacement": "underHeading",
            "linkNearHeading": "カルテ保護の要点",
        },
        "link_shape_note": "見出し「カルテ保護の要点」下に「本院案内」図形（URLリンク未設定）",
    },
    {
        "setNo": 3,
        "fn_heading": "ライン安全の考え方",
        "bib_heading": "参考資料",
        "task8_1": {
            "subtitle": "－生産ラインの守り方－",
            "tocStyle": "自動作成の目次2",
            "tocPlacement": "endOfPage1",
            "insertInShape": False,
            "placementDescription": "1ページ目の最後",
        },
        "task8_2": {"heading": "ライン安全の考え方", "footnoteAnchor": "制御端末", "footnoteText": "ライン停止を防ぐ仕組み"},
        "task8_3": {"footnoteStartNumber": "ⅰ", "numberFormat": "Roman", "numberLabel": "ⅰ"},
        "task8_4": {"heading": "参考資料", "specialChar": "¶", "specialCharName": "¶（段落記号）"},
        "task8_5": {"page": "last", "smartArtType": "階層構造", "smartArtLabels": ["設備", "人", "手順"]},
        "task8_6": {"smartArtColor": "アクセント2", "smartArtAccent": ""},
        "task8_7": {
            "linkShapeText": "文末",
            "linkType": "documentEnd",
            "linkShapePlacement": "underHeading",
            "linkNearHeading": "対策の基本",
        },
        "link_shape_note": "見出し「対策の基本」下に「文末」図形（文末リンク未設定）",
    },
    {
        "setNo": 4,
        "fn_heading": "予約情報の守り方",
        "bib_heading": "関連資料",
        "task8_1": {
            "subtitle": "－予約システムの守り方－",
            "tocStyle": "自動作成の目次2",
            "tocPlacement": "aboveHeading",
            "tocAnchorHeading": "対策の基本",
            "insertInShape": False,
            "placementDescription": "見出し「対策の基本」の上",
        },
        "task8_2": {"heading": "予約情報の守り方", "footnoteAnchor": "予約端末", "footnoteText": "個人情報を保護する手順"},
        "task8_3": {"footnoteStartNumber": "※", "numberFormat": "Symbol", "numberLabel": "※"},
        "task8_4": {"heading": "関連資料", "specialChar": "®", "specialCharName": "®（登録商標記号）"},
        "task8_5": {"page": "last", "smartArtType": "関係", "smartArtLabels": ["予約", "決済", "案内"]},
        "task8_6": {"smartArtColor": "アクセント3", "smartArtAccent": ""},
        "task8_7": {
            "linkShapeText": "先頭へ",
            "linkType": "documentTop",
            "linkShapePlacement": "underHeading",
            "linkNearHeading": "対策の基本",
        },
        "link_shape_note": "見出し「対策の基本」下に「先頭へ」図形（文頭リンク未設定）",
    },
    {
        "setNo": 5,
        "fn_heading": "学習データの保護",
        "bib_heading": "出典一覧",
        "task8_1": {
            "subtitle": "－学習端末の守り方－",
            "tocStyle": "自動作成の目次2",
            "tocPlacement": "beforeFirstHeading1",
            "insertInShape": False,
            "placementDescription": "最初の見出し1の直前",
        },
        "task8_2": {"heading": "学習データの保護", "footnoteAnchor": "学習端末", "footnoteText": "成績情報を守る仕組み"},
        "task8_3": {"footnoteStartNumber": "ア", "numberFormat": "Aiueo", "numberLabel": "ア"},
        "task8_4": {"heading": "出典一覧", "specialChar": "⋯", "specialCharName": "⋯（中点リーダ）"},
        "task8_5": {"page": "last", "smartArtType": "ピラミッド", "smartArtLabels": ["授業", "自習", "共有"]},
        "task8_6": {"smartArtColor": "アクセント4", "smartArtAccent": ""},
        "task8_7": {
            "linkShapeText": "出典へ",
            "linkType": "heading",
            "linkTargetHeading": "出典一覧",
            "linkShapePlacement": "underHeading",
            "linkNearHeading": "学習データの保護",
        },
        "link_shape_note": "見出し「学習データの保護」下に「出典へ」図形（見出しリンク未設定）",
    },
]

OLD_FN_HEADINGS = [
    "POSセキュリティって何？",
    "カルテ保護って何？",
    "ライン安全って何？",
    "予約管理って何？",
    "学習端末って何？",
]


def _task8_1_text(p1: dict) -> str:
    desc = p1["placementDescription"]
    if p1["tocPlacement"] == "underHeading":
        return f"タスク8-1　見出し「{p1['tocAnchorHeading']}」の下に「{p1['tocStyle']}」の目次を挿入します。"
    if p1["tocPlacement"] == "aboveHeading":
        return f"タスク8-1　見出し「{p1['tocAnchorHeading']}」の上に「{p1['tocStyle']}」の目次を挿入します。"
    if p1["tocPlacement"] == "endOfPage1":
        return f"タスク8-1　1ページ目の最後に「{p1['tocStyle']}」の目次を挿入します。"
    if p1["tocPlacement"] == "beforeFirstHeading1":
        return f"タスク8-1　最初の見出し1の直前に「{p1['tocStyle']}」の目次を挿入します。"
    return f"タスク8-1　副題「{p1['subtitle']}」の直後に「{p1['tocStyle']}」の目次を挿入します。"


def _task8_3_text(p3: dict) -> str:
    return f"タスク8-3　脚注の開始番号を「{p3['numberLabel']}」に変更します。"


def _task8_4_text(p4: dict) -> str:
    return f"タスク8-4　見出し「{p4['heading']}」の先頭に「{p4['specialCharName']}」の特殊文字を挿入します。"


def _task8_5_text(p5: dict) -> str:
    labels = "」、「".join(f'"{x}"' for x in p5["smartArtLabels"])
    return (
        f"タスク8-5　最後のページにSmart Art「{p5['smartArtType']}」を挿入し、"
        f"「{labels}」と入力します。文字列の順序は問いません。"
    )


def _task8_6_text(p6: dict) -> str:
    if p6["smartArtColor"] == "モノクロ":
        return f"タスク8-6　最後のページのSmart Artの色を「{p6['smartArtColor']}」に変更します。"
    return f"タスク8-6　最後のページのSmart Artの色を「{p6['smartArtColor']}」に変更します。"


def _task8_7_text(p7: dict) -> str:
    near = p7["linkNearHeading"]
    shape = p7["linkShapeText"]
    if p7["linkType"] == "url":
        return (
            f"タスク8-7　見出し「{near}」の下の「{shape}」と書かれた図形に、"
            f"「{p7['linkUrl']}」のリンクを挿入します。"
        )
    if p7["linkType"] == "heading":
        return (
            f"タスク8-7　見出し「{near}」の下の「{shape}」と書かれた図形に、"
            f"見出し「{p7['linkTargetHeading']}」へのリンクを挿入します。"
        )
    if p7["linkType"] == "documentEnd":
        return f"タスク8-7　見出し「{near}」の下の「{shape}」と書かれた図形に、文末へのリンクを挿入します。"
    return f"タスク8-7　見出し「{near}」の下の「{shape}」と書かれた図形に、文頭に戻るリンクを挿入します。"


def _rekey_body_headings(cb: dict, old_h: str, new_h: str) -> None:
    body = cb.get("bodyUnderHeadings", {})
    if old_h in body:
        body[new_h] = body.pop(old_h)
    elif new_h not in body and body:
        first_key = next(iter(body))
        body[new_h] = body.pop(first_key)


def patch(data: dict) -> dict:
    rules = data["variantRules"]
    rules["scope"] = (
        "MOSスペシャリスト範疇。8-1=目次挿入（図形内・副題下以外）、8-2=脚注、8-3=脚注番号形式、"
        "8-4=特殊文字、8-5=SmartArt、8-6=SmartArt色、8-7=リンク（最下部以外）の操作種別を維持。"
    )
    rules["perSetMustVary"] = [
        "副題と目次配置",
        "脚注見出しと脚注番号形式",
        "文献見出しと特殊文字",
        "SmartArt種類とラベル",
        "SmartArt色",
        "リンク種別と図形位置",
    ]
    rules["mustNotReuseAcrossSets"] = [
        "subtitle",
        "tocPlacement",
        "footnoteAnchor",
        "footnoteStartNumber",
        "specialChar",
        "smartArtType",
        "smartArtLabels",
        "smartArtColor",
        "linkType",
        "linkShapeText",
        "linkUrl",
    ]

    patch_by_set = {p["setNo"]: p for p in SET_PATCHES}
    for s in data["sets"]:
        p = patch_by_set[s["setNo"]]
        old_fn = s["layout"]["headings"][0]
        new_fn = p["fn_heading"]
        bib = p["bib_heading"]
        s["layout"]["headings"][0] = new_fn
        s["layout"]["headings"][-1] = bib

        cb = s["contentBlocks"]
        cb["subtitle"] = p["task8_1"]["subtitle"]
        if "subtitleShape" in cb:
            del cb["subtitleShape"]
        _rekey_body_headings(cb, old_fn, new_fn)
        if old_fn in OLD_FN_HEADINGS:
            _rekey_body_headings(cb, old_fn, new_fn)

        p7 = dict(p["task8_7"])
        cb["linkShape"] = {
            "text": p7["linkShapeText"],
            "placement": p7["linkShapePlacement"],
            "nearHeading": p7["linkNearHeading"],
            "linkType": p7["linkType"],
            "note": p["link_shape_note"],
        }
        if p7.get("linkUrl"):
            cb["linkShape"]["linkUrl"] = p7["linkUrl"]
        if p7.get("linkTargetHeading"):
            cb["linkShape"]["targetHeading"] = p7["linkTargetHeading"]
        cb.pop("bottomShapeText", None)

        p1, p2, p3, p4, p5, p6 = (
            p["task8_1"],
            p["task8_2"],
            p["task8_3"],
            p["task8_4"],
            p["task8_5"],
            p["task8_6"],
        )
        s["taskParams"] = {
            "task8_1": dict(p1),
            "task8_2": dict(p2),
            "task8_3": dict(p3),
            "task8_4": dict(p4),
            "task8_5": dict(p5),
            "task8_6": dict(p6),
            "task8_7": dict(p7),
        }
        s["tasks"] = [
            _task8_1_text(p1),
            f"タスク8-2　見出し「{p2['heading']}」の下にある「{p2['footnoteAnchor']}」の後ろに「\"{p2['footnoteText']}\"」と脚注を挿入します。",
            _task8_3_text(p3),
            _task8_4_text(p4),
            _task8_5_text(p5),
            _task8_6_text(p6),
            _task8_7_text(p7),
        ]

    return apply_outlines_to_json(data)


def main() -> None:
    data = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    data = patch(data)
    JSON_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"patched: {JSON_PATH.name}")


if __name__ == "__main__":
    main()
