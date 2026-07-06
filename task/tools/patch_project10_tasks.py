"""Patch Project10 Word variant JSON: diversify tasks 10-1..10-5 away from textbook patterns."""

from __future__ import annotations

import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
JSON_PATH = BASE / "類題Json" / "MOS_Word類題_project10_配置別_5セット_問題文.json"
TOOLS = Path(__file__).resolve().parent
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
from word_document_outline import apply_outlines_to_json

SET_PATCHES = [
    {
        "setNo": 1,
        "task10_1": {"searchText": "単価", "replaceText": "税込み価格", "occurrenceCount": 3},
        "bodyWithSearchTerms": [
            "メニューの単価は毎月見直します。",
            "単価改定時はPOPを更新してください。",
            "キャンペーン前に単価を再確認します。",
            "端末のパスワードを共有しないこと。",
            "共有端末では個人フォルダを使わない。",
        ],
        "formatTarget": "共有",
        "task10_3": {"bulletFont": "Wingdings", "bulletCharCode": "153"},
        "task10_4": {"formatType": "bold"},
        "task10_5": {"paperSize": "A4", "lineSpacingType": "1.5行", "lineSpacingValue": None},
        "lastTwoLines": [
            "以上がPOSセキュリティの要点です。",
            "全スタッフで共有し実践してください。",
        ],
    },
    {
        "setNo": 2,
        "task10_1": {"searchText": "処方", "replaceText": "電子処方", "occurrenceCount": 3},
        "bodyWithSearchTerms": [
            "外来では処方内容を二重確認します。",
            "処方箋の電子化で転記ミスを減らします。",
            "処方変更時は患者へ説明を行います。",
            "記録の確認は診療終了後に実施します。",
            "薬剤の確認は調剤前にも行います。",
        ],
        "formatTarget": "確認",
        "task10_3": {"bulletFont": "Symbol", "bulletCharCode": "183"},
        "task10_4": {"formatType": "color", "formatColor": "赤"},
        "task10_5": {"paperSize": "A5", "lineSpacingType": "2行", "lineSpacingValue": None},
        "lastTwoLines": [
            "以上がカルテ保護の要点です。",
            "全職員で確認し実践してください。",
        ],
    },
    {
        "setNo": 3,
        "task10_1": {"searchText": "検査", "replaceText": "品質検査", "occurrenceCount": 3},
        "bodyWithSearchTerms": [
            "出荷前の検査は品質を保証します。",
            "検査基準は月次で見直します。",
            "検査記録は三年間保管します。",
            "異常時の手順を全員が把握します。",
            "手順書は現場に掲示します。",
        ],
        "formatTarget": "手順",
        "task10_3": {"bulletFont": "Marlett", "bulletCharCode": "252"},
        "task10_4": {"formatType": "fontSize", "formatSizePt": 14},
        "task10_5": {"paperSize": "レター", "lineSpacingType": "最小値", "lineSpacingValue": 12},
        "lastTwoLines": [
            "以上がライン安全の要点です。",
            "全員で手順を徹底してください。",
        ],
    },
    {
        "setNo": 4,
        "task10_1": {"searchText": "座席", "replaceText": "指定座席", "occurrenceCount": 3},
        "bodyWithSearchTerms": [
            "予約時に座席を確保します。",
            "座席変更は出発24時間前まで受付ます。",
            "座席表は搭乗前に再確認します。",
            "乗客への案内は丁寧に行います。",
            "緊急時の案内は定型文を使用します。",
        ],
        "formatTarget": "案内",
        "task10_3": {"bulletFont": "Wingdings", "bulletCharCode": "113"},
        "task10_4": {"formatType": "color", "formatColor": "濃い青"},
        "task10_5": {"paperSize": "エグゼクティブ", "lineSpacingType": "固定値", "lineSpacingValue": 18},
        "lastTwoLines": [
            "以上が予約管理の要点です。",
            "全員で案内を統一してください。",
        ],
    },
    {
        "setNo": 5,
        "task10_1": {"searchText": "課題", "replaceText": "演習課題", "occurrenceCount": 3},
        "bodyWithSearchTerms": [
            "週次の課題は期限内に提出します。",
            "課題の配布は学習システムで行います。",
            "課題の採点結果は個別に返却します。",
            "提出期限は授業開始前に周知します。",
            "再提出は担任の確認後に認めます。",
        ],
        "formatTarget": "提出",
        "task10_3": {"bulletFont": "Symbol", "bulletCharCode": "167"},
        "task10_4": {"formatType": "color", "formatColor": "紫"},
        "task10_5": {"paperSize": "はがき", "lineSpacingType": "倍数", "lineSpacingValue": 1.25},
        "lastTwoLines": [
            "以上が学習データ保護の要点です。",
            "全生徒に提出ルールを周知してください。",
        ],
    },
]


def _task10_1_text(p1: dict) -> str:
    return (
        f"タスク10-1　文書内のすべての「\"{p1['searchText']}\"」を"
        f"「\"{p1['replaceText']}\"」に置き換えます。"
    )


def _task10_3_text(p3: dict) -> str:
    return (
        f"タスク10-3　見出し「{p3['heading']}」の下の行頭文字を"
        f"フォント「{p3['bulletFont']}」の文字コード「\"{p3['bulletCharCode']}\"」にします。"
    )


def _task10_4_text(p4: dict) -> str:
    target = p4["formatTarget"]
    ft = p4["formatType"]
    if ft == "bold":
        return f"タスク10-4　置換機能を使用して、文書内の全て「\"{target}\"」を太字に変更します。"
    if ft == "color":
        return (
            f"タスク10-4　置換機能を使用して、文書内の全て「\"{target}\"」の"
            f"フォントの色を「{p4['formatColor']}」に変更します。"
        )
    return (
        f"タスク10-4　置換機能を使用して、文書内の全て「\"{target}\"」の"
        f"文字のサイズを「\"{p4['formatSizePt']}\"」ptに変更します。"
    )


def _task10_5_text(p5: dict) -> str:
    paper = p5["paperSize"]
    lst = p5["lineSpacingType"]
    val = p5.get("lineSpacingValue")
    if lst in ("1行", "1.5行", "2行"):
        return f"タスク10-5　文書の用紙サイズを{paper}に変更し、最後の2行の行間を「{lst}」にします。"
    if lst == "倍数":
        return (
            f"タスク10-5　文書の用紙サイズを{paper}に変更し、"
            f"最後の2行の行間を「倍数」の「\"{val}\"」にします。"
        )
    return (
        f"タスク10-5　文書の用紙サイズを{paper}に変更し、"
        f"最後の2行の行間を「{lst}」の「\"{val}\"」ptにします。"
    )


def _format_note(p4: dict) -> str:
    ft = p4["formatType"]
    target = p4["formatTarget"]
    if ft == "bold":
        return f"10-4: 「{target}」は太字未設定"
    if ft == "color":
        return f"10-4: 「{target}」の色は{p4['formatColor']}未設定"
    return f"10-4: 「{target}」のサイズは{p4['formatSizePt']}pt未設定"


def _line_spacing_note(p5: dict) -> str:
    lst = p5["lineSpacingType"]
    val = p5.get("lineSpacingValue")
    if lst in ("1行", "1.5行", "2行"):
        return lst
    if lst == "倍数":
        return f"倍数{val}"
    return f"{lst}{val}pt"


def patch(data: dict) -> dict:
    rules = data["variantRules"]
    rules["scope"] = (
        "MOSスペシャリスト範疇。10-1=一括置換（置換前後は独立）、10-2=画像行頭文字、"
        "10-3=記号フォント行頭文字（Webdings・120以外）、10-4=置換で書式（太字/色/サイズ）、"
        "10-5=用紙サイズ（B5以外）＋末尾2行の行間（1.6行以外）の操作種別を維持。"
    )
    rules["perSetMustVary"] = [
        "置換ペア",
        "10-2見出しと画像名",
        "10-3フォントと文字コード",
        "書式変更対象と種別",
        "用紙サイズと行間",
    ]
    rules["mustNotReuseAcrossSets"] = [
        "searchText",
        "replaceText",
        "heading10_2",
        "bulletImageName",
        "heading10_3",
        "formatTarget",
        "bulletFont",
        "bulletCharCode",
        "paperSize",
    ]

    patch_by_set = {p["setNo"]: p for p in SET_PATCHES}
    for s in data["sets"]:
        p = patch_by_set[s["setNo"]]
        cb = s["contentBlocks"]
        p1 = {**s["taskParams"]["task10_1"], **p["task10_1"]}
        p2 = dict(s["taskParams"]["task10_2"])
        p3 = {**s["taskParams"]["task10_3"], **p["task10_3"]}
        p4 = {
            "formatTarget": p["formatTarget"],
            "operation": "replaceWithFormat",
            **p["task10_4"],
        }
        p5 = {**s["taskParams"]["task10_5"], **p["task10_5"], "targetLines": "lastTwo"}
        p5.pop("lineSpacing", None)

        cb["bodyWithSearchTerms"] = p["bodyWithSearchTerms"]
        cb["searchOccurrences"] = {
            p1["searchText"]: p1["occurrenceCount"],
            p4["formatTarget"]: sum(
                line.count(p4["formatTarget"]) for line in p["bodyWithSearchTerms"]
            )
            + sum(line.count(p4["formatTarget"]) for line in p["lastTwoLines"]),
        }
        cb["lastTwoLines"] = p["lastTwoLines"]

        s["taskParams"] = {
            "task10_1": p1,
            "task10_2": p2,
            "task10_3": p3,
            "task10_4": p4,
            "task10_5": p5,
        }
        s["tasks"] = [
            _task10_1_text(p1),
            f"タスク10-2　見出し「{p2['heading']}」の下の段落番号を"
            f"「{p2['bulletImageName']}」の画像の行頭文字に変更します。",
            _task10_3_text(p3),
            _task10_4_text(p4),
            _task10_5_text(p5),
        ]

        outline = cb.get("documentOutline", {})
        for sec in outline.get("sections", []):
            note = sec.get("note", "")
            if note.startswith("10-1:"):
                sec["note"] = f"10-1: 「{p1['searchText']}」は置換前のまま複数回出現"
            elif note.startswith("10-3:"):
                sec["note"] = (
                    f"10-3: {p3['bulletFont']} {p3['bulletCharCode']}は未設定"
                )
            elif note.startswith("10-5:"):
                ls = _line_spacing_note(p5)
                sec["note"] = f"10-5: 文書末尾2行は行間{ls}未設定"
                for child in sec.get("children", []):
                    _update_p10_child_notes(child, p5, p4)

        if outline.get("sections"):
            last_sec = outline["sections"][-1]
            if last_sec.get("text") == "まとめ":
                fmt = _format_note(p4)
                ls = _line_spacing_note(p5)
                last_sec["note"] = f"{fmt} / 10-5: 末尾2行行間{ls}未設定"

    return apply_outlines_to_json(data)


def _update_p10_child_notes(node: dict, p5: dict, p4: dict) -> None:
    for child in node.get("children", []):
        body = child.get("body", [])
        if body and "用紙B5" in body[-1]:
            body[-1] = (
                f"用紙{p5['paperSize']}・行間{_line_spacing_note(p5)}は未設定（10-5）"
            )
        _update_p10_child_notes(child, p5, p4)


def main() -> None:
    data = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    data = patch(data)
    JSON_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"patched: {JSON_PATH.name}")


if __name__ == "__main__":
    main()
