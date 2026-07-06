"""Patch Project9 Word variant JSON: diversify tasks 9-1..9-5 away from textbook patterns."""

from __future__ import annotations

import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
JSON_PATH = BASE / "類題Json" / "MOS_Word類題_project9_配置別_5セット_問題文.json"
TOOLS = Path(__file__).resolve().parent
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
from word_document_outline import apply_outlines_to_json

SORT_ORDER_LABELS = {
    "desc": "降順",
    "asc": "昇順",
    "alpha": "あいうえお順",
}

MARGIN_SIDE_LABELS = {
    "left": "左",
    "top": "上",
    "bottom": "下",
    "right": "右",
}

SEPARATOR_OPS = {
    "comma": ("convertTableToCommaText", "コンマ区切り"),
    "tab": ("convertTableToTabText", "タブ区切り"),
    "paragraph": ("convertTableToParagraphText", "段落区切り"),
}

TRACK_CHANGE_OPS = {
    "acceptAll": "文書内の変更履歴をすべて承認してください。",
    "revertAll": "文書内の変更履歴をすべて修正前の状態に戻してください。",
    "rejectAllAndStop": "文書内の変更履歴をすべて取り消して変更履歴の記録を停止します。",
}

SET_PATCHES = [
    {
        "setNo": 1,
        "restartAt": 2,
        "separator": "comma",
        "marginSide": "left",
        "marginMm": 3,
        "sortColumn": "商品",
        "sortOrder": "alpha",
        "trackChangesOp": "acceptAll",
    },
    {
        "setNo": 2,
        "restartAt": 3,
        "separator": "tab",
        "marginSide": "top",
        "marginMm": 5,
        "sortColumn": "5月",
        "sortOrder": "asc",
        "trackChangesOp": "revertAll",
    },
    {
        "setNo": 3,
        "restartAt": 5,
        "separator": "paragraph",
        "marginSide": "bottom",
        "marginMm": 2,
        "sortColumn": "製品",
        "sortOrder": "alpha",
        "trackChangesOp": "rejectAllAndStop",
    },
    {
        "setNo": 4,
        "restartAt": 10,
        "separator": "comma",
        "marginSide": "left",
        "marginMm": 6,
        "sortColumn": "4月",
        "sortOrder": "desc",
        "trackChangesOp": "acceptAll",
    },
    {
        "setNo": 5,
        "restartAt": 7,
        "separator": "tab",
        "marginSide": "top",
        "marginMm": 8,
        "sortColumn": "講座",
        "sortOrder": "alpha",
        "trackChangesOp": "revertAll",
    },
]


def _task9_1_text(p1: dict) -> str:
    return (
        f"タスク9-1　見出し「{p1['heading']}」の下にある「{p1['numberedParagraph']}」の"
        f"段落番号を「{p1['restartAt']}」から開始するようにします。"
    )


def _task9_2_text(p2: dict) -> str:
    return f"タスク9-2　見出し「{p2['heading']}」の下の表を、{p2['separatorLabel']}の文字列に変更します。"


def _task9_3_text(p3: dict) -> str:
    side = MARGIN_SIDE_LABELS[p3["marginSide"]]
    return (
        f"タスク9-3　「{p3['tableHeading']}」の下の表のセルの{side}の余白を"
        f"「\"{p3['marginMm']}\"mm」にします。"
    )


def _task9_4_text(p4: dict) -> str:
    order = SORT_ORDER_LABELS[p4["sortOrder"]]
    return (
        f"タスク9-4　「{p4['tableHeading']}」の下の表を、"
        f"「{p4['sortColumn']}」の{order}に並べ替えます。"
    )


def _task9_5_text(p5: dict) -> str:
    return f"タスク9-5　{TRACK_CHANGE_OPS[p5['operation']]}"


def patch(data: dict) -> dict:
    rules = data["variantRules"]
    rules["scope"] = (
        "MOSスペシャリスト範疇。9-1=段落番号再開（1以外）、9-2=表を文字列変換（コンマは2セットのみ）、"
        "9-3=セル余白（右・4mm以外）、9-4=表並べ替え（合計列・降順以外も）、"
        "9-5=変更履歴操作（承認のみ/修正前に戻す/取り消し＋記録停止）、9-6=変更履歴ロックの操作種別を維持。"
    )
    rules["perSetMustVary"] = [
        "段落番号の再開値",
        "表の文字列変換区切り",
        "セル余白の辺と数値",
        "並べ替え列と順序",
        "変更履歴の操作種別",
        "変更履歴パスワード",
    ]
    rules["mustNotReuseAcrossSets"] = [
        "numberedParagraph",
        "table9_2Heading",
        "table9_3Heading",
        "sortColumn",
        "trackChangesPassword",
        "restartAt",
        "separator",
        "marginSide",
        "marginMm",
        "trackChangesOp",
    ]

    patch_by_set = {p["setNo"]: p for p in SET_PATCHES}
    for s in data["sets"]:
        p = patch_by_set[s["setNo"]]
        tp = s["taskParams"]
        op, sep_label = SEPARATOR_OPS[p["separator"]]

        p1 = {**tp["task9_1"], "restartAt": p["restartAt"]}
        p2 = {
            **tp["task9_2"],
            "operation": op,
            "separator": p["separator"],
            "separatorLabel": sep_label,
        }
        p3 = {
            **tp["task9_3"],
            "marginSide": p["marginSide"],
            "marginSideLabel": MARGIN_SIDE_LABELS[p["marginSide"]],
            "marginMm": p["marginMm"],
        }
        p4 = {
            **tp["task9_4"],
            "sortColumn": p["sortColumn"],
            "sortOrder": p["sortOrder"],
            "sortOrderLabel": SORT_ORDER_LABELS[p["sortOrder"]],
        }
        p5 = {"operation": p["trackChangesOp"]}
        p6 = dict(tp["task9_6"])

        s["taskParams"] = {
            "task9_1": p1,
            "task9_2": p2,
            "task9_3": p3,
            "task9_4": p4,
            "task9_5": p5,
            "task9_6": p6,
        }
        s["tasks"] = [
            _task9_1_text(p1),
            _task9_2_text(p2),
            _task9_3_text(p3),
            _task9_4_text(p4),
            _task9_5_text(p5),
            f"タスク9-6　他のユーザーが変更履歴を編集できないようにロックします。"
            f"パスワードは「\"{p6['trackChangesPassword']}\"」にします。",
        ]

    return apply_outlines_to_json(data)


def main() -> None:
    data = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    data = patch(data)
    JSON_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"patched: {JSON_PATH.name}")


if __name__ == "__main__":
    main()
