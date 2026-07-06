"""Patch Project7 Word variant JSON: 7-2 property (not company), 7-3~7-6 header/footer, 7-7/7-8 save."""
import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
JSON_PATH = BASE / "類題Json" / "MOS_Word類題_project7_配置別_5セット_問題文.json"
APP_PATH = BASE / "類題Json" / "MOS_Word類題_project7_配置別_5セット_問題文_アプリ用.json"

SET_PATCHES = [
    {
        "setNo": 1,
        "theme": "カフェ",
        "task7_2": {"propertyName": "作成者", "propertyKey": "Author", "propertyValue": "焙煎カフェ係"},
        "task7_3": {"headerText": "コーヒー焙煎", "headerType": "primary"},
        "task7_4": {"headerText": "体験会案内", "headerType": "firstPage"},
        "task7_5": {"footerType": "pageNumber", "operation": "insertPageNumber"},
        "task7_6": {"footerText": "焙煎カフェ本店", "footerType": "primary"},
        "task7_7": {"txtBaseName": "焙煎体験会", "saveFormat": "txt"},
        "task7_8": {"docmBaseName": "焙煎体験会", "readPassword": "cf26", "saveFormat": "docm"},
    },
    {
        "setNo": 2,
        "theme": "医療",
        "task7_2": {"propertyName": "件名", "propertyKey": "Subject", "propertyValue": "健康講座のご案内"},
        "task7_3": {"headerText": "健診案内", "headerType": "primary"},
        "task7_4": {"headerText": "本院健診係", "headerType": "firstPage"},
        "task7_5": {"footerType": "pageNumber", "operation": "insertPageNumber"},
        "task7_6": {"footerText": "メディカルセンター本院", "footerType": "primary"},
        "task7_7": {"txtBaseName": "健康講座案内", "saveFormat": "txt"},
        "task7_8": {"docmBaseName": "健康講座案内", "readPassword": "hm26", "saveFormat": "docm"},
    },
    {
        "setNo": 3,
        "theme": "製造",
        "task7_2": {"propertyName": "管理者", "propertyKey": "Manager", "propertyValue": "第一工場安全係"},
        "task7_3": {"headerText": "品質管理", "headerType": "primary"},
        "task7_4": {"headerText": "安全講習係", "headerType": "firstPage"},
        "task7_5": {"footerType": "pageNumber", "operation": "insertPageNumber"},
        "task7_6": {"footerText": "第一工場 総務部", "footerType": "primary"},
        "task7_7": {"txtBaseName": "安全講習会", "saveFormat": "txt"},
        "task7_8": {"docmBaseName": "安全講習会", "readPassword": "mf26", "saveFormat": "docm"},
    },
    {
        "setNo": 4,
        "theme": "旅行",
        "task7_2": {"propertyName": "タイトル", "propertyKey": "Title", "propertyValue": "添乗員研修案内"},
        "task7_3": {"headerText": "旅程案内", "headerType": "primary"},
        "task7_4": {"headerText": "研修係", "headerType": "firstPage"},
        "task7_5": {"footerType": "pageNumber", "operation": "insertPageNumber"},
        "task7_6": {"footerText": "トラベル研修センター", "footerType": "primary"},
        "task7_7": {"txtBaseName": "添乗員研修", "saveFormat": "txt"},
        "task7_8": {"docmBaseName": "添乗員研修", "readPassword": "tr26", "saveFormat": "docm"},
    },
    {
        "setNo": 5,
        "theme": "学習塾",
        "task7_2": {"propertyName": "分類", "propertyKey": "Category", "propertyValue": "保護者説明資料"},
        "task7_3": {"headerText": "学習計画", "headerType": "primary"},
        "task7_4": {"headerText": "進路指導係", "headerType": "firstPage"},
        "task7_5": {"footerType": "pageNumber", "operation": "insertPageNumber"},
        "task7_6": {"footerText": "エデュ学習塾本部", "footerType": "primary"},
        "task7_7": {"txtBaseName": "保護者説明会", "saveFormat": "txt"},
        "task7_8": {"docmBaseName": "保護者説明会", "readPassword": "ed26", "saveFormat": "docm"},
    },
]


def _task_texts(p: dict) -> list[str]:
    p2, p3, p4, p5, p6, p7, p8 = (
        p["task7_2"],
        p["task7_3"],
        p["task7_4"],
        p["task7_5"],
        p["task7_6"],
        p["task7_7"],
        p["task7_8"],
    )
    return [
        "タスク7-1　文書の互換モードを解除します。メッセージが表示された場合は「OK」をクリックします。",
        f"タスク7-2　文章のプロパティの{p2['propertyName']}に「\"{p2['propertyValue']}\"」と設定します。",
        f"タスク7-3　文書に「{p3['headerText']}」のヘッダーを挿入します。",
        f"タスク7-4　最初のページのみ、ヘッダーに「{p4['headerText']}」と入力します。",
        "タスク7-5　フッターにページ番号を挿入します。",
        f"タスク7-6　フッターに「{p6['footerText']}」と入力します。",
        f"タスク7-7　文書に「\"{p7['txtBaseName']}\"」という名前を付けてテキストファイルとして保存します。ファイルの変換は既定値のままにします。",
        f"タスク7-8　この文書のコピーをマクロ有効文書として保存します。「名前を付けて保存」画面で読み取りパスワードを「\"{p8['readPassword']}\"」に設定すること。",
    ]


def _update_outline_note(s: dict) -> None:
    outline = s.get("contentBlocks", {}).get("documentOutline")
    if not outline:
        return
    for sec in outline.get("sections", []):
        if sec.get("note") and "7-2" in sec.get("note", ""):
            sec["note"] = "プロパティ・ヘッダー・フッターは未設定（7-2〜7-6は受験者操作）"


def patch(data: dict) -> dict:
    rules = data["variantRules"]
    rules["scope"] = (
        "MOSスペシャリスト範疇。7-1=互換モード解除、7-2=文書プロパティ（会社名以外）、"
        "7-3=ヘッダー挿入、7-4=最初のページのヘッダー、7-5=フッターページ番号、7-6=フッターテキスト、"
        "7-7=txt保存、7-8=docm保存＋読み取りパスワードの操作種別を維持。"
    )
    rules["perSetMustVary"] = [
        "導入文と本文テーマ",
        "プロパティ項目と値",
        "ヘッダー文字列",
        "フッター文字列",
        "保存ベース名",
        "読み取りパスワード",
    ]
    rules["mustNotReuseAcrossSets"] = [
        "propertyValue",
        "headerText",
        "firstPageHeaderText",
        "footerText",
        "txtBaseName",
        "readPassword",
    ]
    rules["specificity"] = "プロパティ・ヘッダー・フッター・保存名・パスワードを完全一致で指定"

    ui = data["wordUiNames"]
    ui["footer"] = "挿入 → ヘッダーとフッター → フッター"
    ui["documentPropertyFields"] = "ファイル → 情報 → プロパティ → 詳細プロパティの表示"

    patch_by_set = {p["setNo"]: p for p in SET_PATCHES}
    for s in data["sets"]:
        p = patch_by_set[s["setNo"]]
        s["tasks"] = _task_texts(p)
        s["taskParams"] = {
            "task7_1": {"operation": "upgradeDocument"},
            "task7_2": dict(p["task7_2"]),
            "task7_3": dict(p["task7_3"]),
            "task7_4": dict(p["task7_4"]),
            "task7_5": dict(p["task7_5"]),
            "task7_6": dict(p["task7_6"]),
            "task7_7": dict(p["task7_7"]),
            "task7_8": dict(p["task7_8"]),
        }
        _update_outline_note(s)
    return data


def write_app_json(data: dict) -> dict:
    import re

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


def main() -> None:
    data = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    data = patch(data)
    JSON_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    app = write_app_json(data)
    APP_PATH.write_text(json.dumps(app, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"patched: {JSON_PATH.name}")
    print(f"patched: {APP_PATH.name}")
    for s in data["sets"]:
        print(f"  set{s['setNo']}: {len(s['tasks'])} tasks")


if __name__ == "__main__":
    main()
