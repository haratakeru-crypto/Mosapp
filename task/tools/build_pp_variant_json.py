"""Author and build PowerPoint variant JSON for Projects 1-6 (5 sets each)."""

from __future__ import annotations

import json
import sys
from copy import deepcopy
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
TOOLS = Path(__file__).resolve().parent
if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
from pp_projects_7_10 import (
    FORBIDDEN_P10,
    FORBIDDEN_P7,
    FORBIDDEN_P8,
    FORBIDDEN_P9,
    P10_SETS,
    P7_SETS,
    P8_SETS,
    P9_SETS,
    build_project10_set,
    build_project7_set,
    build_project8_set,
    build_project9_set,
)

THEMES = [
    {
        "setNo": 1,
        "theme": "カフェ",
        "colorTheme": "茶、スライドの背景2",
        "designTheme": "イオン",
        "variant": "オフィス",
        "fontTheme": "メイリオ",
    },
    {
        "setNo": 2,
        "theme": "医療",
        "colorTheme": "緑、アクセント1",
        "designTheme": "センチュリー",
        "variant": "モダン",
        "fontTheme": "メイリオ",
    },
    {
        "setNo": 3,
        "theme": "製造",
        "colorTheme": "青灰、アクセント2",
        "designTheme": "イオン",
        "variant": "シンプル",
        "fontTheme": "游ゴシック",
    },
    {
        "setNo": 4,
        "theme": "旅行",
        "colorTheme": "水色、アクセント5",
        "designTheme": "バーチ",
        "variant": "モダン",
        "fontTheme": "メイリオ",
    },
    {
        "setNo": 5,
        "theme": "学習塾",
        "colorTheme": "オレンジ、アクセント4",
        "designTheme": "ロード",
        "variant": "オフィス",
        "fontTheme": "游ゴシック",
    },
]

FORBIDDEN_P1 = [
    "英語教育", "UP Rabbit", "教育理念", "募集要項", "ご提案のポイント",
    "テーブルスライド", "テキスト１スライド", "テキスト1スライド",
]
FORBIDDEN_P2 = [
    "MOS合格対策", "スプリット", "ワイプアウト", "黒板", "ジャンプしてターン",
    "今年度募集について", "ニュートロン", "男の子",
]
FORBIDDEN_P3 = [
    "Power Point 新機能", "タイムライン", "ターゲットリスト", "虫眼鏡",
    "機能の概要", "画面録画で説明", "ズーム機能で訴求力アップ",
    "デザインアイデアで魅力的", "伝わるスライドの要素",
    "グラデーション循環-アクセント６", "グラデーション 循環-アクセント6",
]
FORBIDDEN_P4 = [
    "教育者必見", "いつでも体験可能です!", "英語教育", "子供",
    "青い図形", "楕円 ぼかし", "テクスチャライザー",
]
FORBIDDEN_P5 = [
    "MOS合格対策", "対策講座", "PC教室", "通信講座",
    "MOSって何？", "男の子", "スマイル",
]
FORBIDDEN_P6 = [
    "伝わる", "書式のポイント", "プレゼンテーションのテクニック",
]

POWERPOINT_UI = {
    "build": "2508",
    "newSlide": "ホーム → 新しいスライド",
    "hideSlide": "スライドショー → 非表示スライド",
    "layout": "ホーム → レイアウト",
    "columns": "ホーム → 段の追加または削除 → 2 段組み",
    "characterSpacing": "ホーム → 文字の間隔 → その他の間隔",
    "section": "ホーム → セクション",
    "zoom": "挿入 → ズーム",
    "transitionTab": "画面切り替え",
    "animationTab": "アニメーション",
    "smartArt": "SmartArt グラフィックの挿入",
    "model3d": "挿入 → 3D モデル",
}


def _slide_no(n: int) -> str:
    return str(n)


# ---------------------------------------------------------------------------
# Project 1
# ---------------------------------------------------------------------------

P1_INTROS = {
    1: "あなたは駅前カフェの新店舗提案資料を作成しています。",
    2: "あなたは本院の健診案内プレゼン資料を作成しています。",
    3: "あなたは工場の安全啓発プレゼン資料を作成しています。",
    4: "あなたは国内ツアー企画の提案資料を作成しています。",
    5: "あなたは秋期講座のご案内プレゼン資料を作成しています。",
}

P1_SETS = [
    {
        "zoomTitle": "提案の要点",
        "zoomLinks": ["1.メニュー方針", "4.開店日程"],
        "zoomExclude": [1, 8],
        "sectionNew": "売上振り返り",
        "sectionRename": "店舗概要",
        "layoutInsert": "比較",
        "layoutChange": "1枚目の項目とテキスト",
        "placeholderText": "オープン直後の来客層",
        "spacingPt": 3,
        "slides": {
            1: ("タイトルスライド", "駅前カフェ 春の集客プラン", "地域密着型メニュー刷新"),
            2: ("タイトルとコンテンツ", "1.メニュー方針", None),
            3: ("タイトルとコンテンツ", "2.来客分析", None),
            4: ("タイトルとコンテンツ", "補足データ", None),
            5: ("タイトルとコンテンツ", "3.オペレーション", None),
            6: ("タイトルとコンテンツ", "4.スタッフ配置", ["朝班と夕班の役割分担", "ピーク時の動線確保", "清掃チェックリストの運用"]),
            7: ("タイトルとコンテンツ", "5.販促施策", None),
            8: ("タイトルとコンテンツ", "4.開店日程", ["試験営業の実施日", "本開店までの準備項目", "地域告知のタイミング"]),
        },
    },
    {
        "zoomTitle": "案内のポイント",
        "zoomLinks": ["1.健診メニュー", "4.予約方法"],
        "zoomExclude": [1, 8],
        "sectionNew": "運用確認",
        "sectionRename": "診療体制",
        "layoutInsert": "コンテンツとキャプション",
        "layoutChange": "1枚目の項目とテキスト",
        "placeholderText": "初めて受診する方の流れ",
        "spacingPt": 5,
        "slides": {
            1: ("タイトルスライド", "本院 健診案内", "生活習慣病の早期発見"),
            2: ("タイトルとコンテンツ", "1.健診メニュー", None),
            3: ("タイトルとコンテンツ", "2.受診の流れ", None),
            4: ("タイトルとコンテンツ", "参考資料", None),
            5: ("タイトルとコンテンツ", "3.注意事項", None),
            6: ("タイトルとコンテンツ", "4.問診の準備", ["前日の食事制限", "服薬の確認事項", "検査結果の受け取り"]),
            7: ("タイトルとコンテンツ", "5.費用補助", None),
            8: ("タイトルとコンテンツ", "4.予約方法", ["Web予約の手順", "電話予約の受付時間", "キャンセルポリシー"]),
        },
    },
    {
        "zoomTitle": "啓発の要点",
        "zoomLinks": ["1.安全基準", "4.点検日程"],
        "zoomExclude": [1, 8],
        "sectionNew": "改善計画",
        "sectionRename": "工場概要",
        "layoutInsert": "2つのコンテンツ",
        "layoutChange": "1枚目の項目とテキスト",
        "placeholderText": "ライン起動前の確認手順",
        "spacingPt": 6,
        "slides": {
            1: ("タイトルスライド", "第一工場 安全啓発", "ライン作業の基本ルール"),
            2: ("タイトルとコンテンツ", "1.安全基準", None),
            3: ("タイトルとコンテンツ", "2.危険予知", None),
            4: ("タイトルとコンテンツ", "作業記録", None),
            5: ("タイトルとコンテンツ", "3.保護具", None),
            6: ("タイトルとコンテンツ", "4.作業手順", ["ロックアウトの実施", "異物混入の防止", "終業時の清掃確認"]),
            7: ("タイトルとコンテンツ", "5.教育計画", None),
            8: ("タイトルとコンテンツ", "4.点検日程", ["週次点検の担当", "月次演習の内容", "記録の保管場所"]),
        },
    },
    {
        "zoomTitle": "企画の要点",
        "zoomLinks": ["1.ツアー概要", "4.申込手順"],
        "zoomExclude": [1, 8],
        "sectionNew": "収益振り返り",
        "sectionRename": "旅行企画",
        "layoutInsert": "比較",
        "layoutChange": "1枚目の項目とテキスト",
        "placeholderText": "初参加のお客様向け案内",
        "spacingPt": 2,
        "slides": {
            1: ("タイトルスライド", "国内線ツアー 夏企画", "家族向け3日間コース"),
            2: ("タイトルとコンテンツ", "1.ツアー概要", None),
            3: ("タイトルとコンテンツ", "2.行程表", None),
            4: ("タイトルとコンテンツ", "添乗員メモ", None),
            5: ("タイトルとコンテンツ", "3.宿泊施設", None),
            6: ("タイトルとコンテンツ", "4.移動手段", ["バス座席の割当", "荷物預かりの案内", "集合時間の周知"]),
            7: ("タイトルとコンテンツ", "5.特典内容", None),
            8: ("タイトルとコンテンツ", "4.申込手順", ["Web申込の流れ", "支払い期限", "変更手数料"]),
        },
    },
    {
        "zoomTitle": "講座の要点",
        "zoomLinks": ["1.講座構成", "4.申込案内"],
        "zoomExclude": [1, 8],
        "sectionNew": "振り返り",
        "sectionRename": "塾概要",
        "layoutInsert": "コンテンツとキャプション",
        "layoutChange": "1枚目の項目とテキスト",
        "placeholderText": "新入生向けの学習習慣",
        "spacingPt": 7,
        "slides": {
            1: ("タイトルスライド", "秋期講座 ご案内", "高校2年生向け特訓"),
            2: ("タイトルとコンテンツ", "1.講座構成", None),
            3: ("タイトルとコンテンツ", "2.カリキュラム", None),
            4: ("タイトルとコンテンツ", "進度表", None),
            5: ("タイトルとコンテンツ", "3.講師紹介", None),
            6: ("タイトルとコンテンツ", "4.学習法", ["週次テストの活用", "質問対応の時間", "自習室の利用ルール"]),
            7: ("タイトルとコンテンツ", "5.合格実績", None),
            8: ("タイトルとコンテンツ", "4.申込案内", ["説明会の日程", "早期申込特典", "兄弟姉妹割引"]),
        },
    },
]


def _p1_slide_map(spec: dict) -> list[dict]:
    sm = []
    for no in range(1, 9):
        layout, title, subtitle = spec["slides"][no]
        entry = {"slideNo": no, "layout": layout, "title": title}
        if subtitle and isinstance(subtitle, list):
            entry["bullets"] = subtitle
        elif subtitle:
            entry["subtitle"] = subtitle
        if no == 1:
            entry["notes"] = "1-7: セクション名は未設定（タイトルなしのセクション）"
        if no == 4:
            entry["notes"] = "1-1: ここに新規スライド挿入前の既存スライド / 1-2: 挿入後に非表示（事前は表示）"
        if no == 5:
            entry["notes"] = "1-3: レイアウト変更・プレースホルダ入力は未実施"
        if no == 6:
            entry["notes"] = "1-4: 2段組み未設定"
        if no == 8:
            entry["notes"] = "1-5: 文字間隔未変更 / 1-6: セクション未追加"
        sm.append(entry)
    return sm


def _p1_tasks(spec: dict, intro: str) -> list[str]:
    zt = spec["zoomTitle"]
    z1, z2 = spec["zoomLinks"]
    lay = spec["layoutInsert"]
    lay2 = spec["layoutChange"]
    ph = spec["placeholderText"]
    sp = spec["spacingPt"]
    sec = spec["sectionNew"]
    sec2 = spec["sectionRename"]
    return [
        f"タスク1-1　スライド４に、レイアウト「{lay}」のスライドを挿入します。",
        "タスク1-2　スライド４を非表示にします。",
        f"タスク1-3　スライド５のレイアウトを「{lay2}」に変更します。"
        f"上側の［マスターテキストのスタイルを編集する］のプレースホルダーに「\"{ph}\"」と入力します。",
        "タスク1-4　スライド６の箇条書きを2段組みに変更します。",
        f"タスク1-5　スライド８の箇条書きのプレースホルダーの文字の間隔を広げます。幅を「\"{sp}\"pt」にします。",
        f"タスク1-6　スライド８にセクションを追加します。セクション名は「\"{sec}\"」にします。",
        f"タスク1-7　スライド１のセクション名を「\"{sec2}\"」とします。",
        f"タスク1-8　スライド１の後ろにサマリーズームスライドを挿入し［{z1}］、［{z2}］の各スライドへのリンクを作成します。"
        f"スライド１と８へのリンクは含めません。タイトルに「\"{zt}\"」と入力します。",
    ]


def _p1_task_params(spec: dict) -> dict:
    return {
        "task1_1": {
            "insertAfterSlide": 3,
            "newSlidePosition": 4,
            "layoutUi": spec["layoutInsert"],
            "operation": "insertSlideWithLayout",
        },
        "task1_2": {"slideNo": 4, "operation": "hideSlide"},
        "task1_3": {
            "slideNo": 5,
            "layoutUi": spec["layoutChange"],
            "placeholderText": spec["placeholderText"],
            "placeholderRole": "マスターテキストのスタイルを編集する",
            "operation": "changeLayoutAndText",
        },
        "task1_4": {"slideNo": 6, "columns": 2, "operation": "twoColumns"},
        "task1_5": {
            "slideNo": 8,
            "spacingPt": spec["spacingPt"],
            "spacingMode": "文字間隔を広げる",
            "operation": "characterSpacing",
        },
        "task1_6": {
            "slideNo": 8,
            "sectionName": spec["sectionNew"],
            "operation": "addSection",
        },
        "task1_7": {
            "slideNo": 1,
            "sectionName": spec["sectionRename"],
            "operation": "renameSection",
        },
        "task1_8": {
            "afterSlide": 1,
            "linkTitles": spec["zoomLinks"],
            "excludeSlideNos": spec["zoomExclude"],
            "zoomTitle": spec["zoomTitle"],
            "operation": "summaryZoom",
        },
    }


def build_project1_set(set_no: int, theme_meta: dict, spec: dict) -> dict:
    intro = P1_INTROS[set_no]
    slide_map = _p1_slide_map(spec)
    return {
        "setNo": set_no,
        "theme": theme_meta["theme"],
        "workbook": f"MOS_PowerPoint類題_project1_配置別_セット{set_no}_{theme_meta['theme']}.pptx",
        "problemStatement": intro,
        "tasks": _p1_tasks(spec, intro),
        "layout": {
            "slideCount": 8,
            "slideSize": "ワイド スクリーン (16:9)",
            "designTheme": theme_meta["designTheme"],
            "variant": theme_meta["variant"],
            "colorTheme": theme_meta["colorTheme"],
            "fontTheme": theme_meta["fontTheme"],
            "sections": [
                {"startSlide": 1, "name": "（タイトルなしのセクション・名称未設定）"},
                {"startSlide": 8, "name": "（1-6実施前はセクション未追加）"},
            ],
        },
        "contentBlocks": {
            "slideMap": slide_map,
            "objects": {},
            "preTaskState": {
                "hiddenSlides": [],
                "slide4Inserted": False,
                "slide4Hidden": False,
                "summaryZoomInserted": False,
                "twoColumnsOnSlide6": False,
                "characterSpacingOnSlide8": False,
            },
        },
        "taskParams": _p1_task_params(spec),
    }


# ---------------------------------------------------------------------------
# Project 2
# ---------------------------------------------------------------------------

P2_INTROS = {
    1: "あなたはカフェのバリスタ研修ガイド資料を作成しています。",
    2: "あなたは医療事務の新人研修資料を作成しています。",
    3: "あなたは製造現場の設備保全研修資料を作成しています。",
    4: "あなたは旅行代理店の新人研修資料を作成しています。",
    5: "あなたは学習塾の講師研修資料を作成しています。",
}

P2_SETS = [
    {
        "transitionAll": "プッシュ", "transitionOption": "右から",
        "durationSec": 2, "perSlideEffect": "フェード", "perSlideNos": [3, 4, 5],
        "autoAdvanceSec": 4, "model3d": "地球儀", "modelAnim": "ターンテーブル",
        "img1": "バリスタ端末", "img2": "確認マーク", "animDir": "右下から", "animDur": 0.75,
        "bulletSlideTitle": "研修スケジュールについて", "bulletEffect": "ひし形",
        "motionShape": "星", "motionPath": "プラス", "circleNoAnim": True,
    },
    {
        "transitionAll": "フェード", "transitionOption": "",
        "durationSec": 4, "perSlideEffect": "プッシュ", "perSlideNos": [3, 4, 5],
        "autoAdvanceSec": 6, "model3d": "聴診器", "modelAnim": "ジャンプ",
        "img1": "受付端末", "img2": "案内マーク", "animDir": "左上から", "animDur": 1.0,
        "bulletSlideTitle": "受付オペレーションについて", "bulletEffect": "チェックマーク",
        "motionShape": "星", "motionPath": "ハート", "circleNoAnim": True,
    },
    {
        "transitionAll": "渦巻き", "transitionOption": "",
        "durationSec": 5, "perSlideEffect": "切り替え", "perSlideNos": [3, 4, 5],
        "autoAdvanceSec": 7, "model3d": "工具箱", "modelAnim": "回転",
        "img1": "保全端末", "img2": "注意マーク", "animDir": "左下から", "animDur": 0.85,
        "bulletSlideTitle": "保全手順の確認について", "bulletEffect": "四角形",
        "motionShape": "星", "motionPath": "ループ", "circleNoAnim": True,
    },
    {
        "transitionAll": "カバー", "transitionOption": "左から",
        "durationSec": 2.5, "perSlideEffect": "渦巻き", "perSlideNos": [3, 4, 5],
        "autoAdvanceSec": 8, "model3d": "飛行機", "modelAnim": "ターンテーブル",
        "img1": "地球儀端末", "img2": "クエスチョンマーク", "animDir": "右上から", "animDur": 1.2,
        "bulletSlideTitle": "予約対応の基本について", "bulletEffect": "ひし形",
        "motionShape": "星", "motionPath": "ベンド", "circleNoAnim": True,
    },
    {
        "transitionAll": "切り替え", "transitionOption": "",
        "durationSec": 4.5, "perSlideEffect": "フェード", "perSlideNos": [3, 4, 5],
        "autoAdvanceSec": 10, "model3d": "ノートパソコン", "modelAnim": "ジャンプ",
        "img1": "ノートPC端末", "img2": "ヒントマーク", "animDir": "右下から", "animDur": 0.9,
        "bulletSlideTitle": "授業運営の基本について", "bulletEffect": "チェックマーク",
        "motionShape": "星", "motionPath": "プラス", "circleNoAnim": True,
    },
]


def _p2_slide_map(spec: dict, set_no: int) -> list[dict]:
    titles = {
        1: ["バリスタ研修ガイド", "研修の目的", "設備の基本", "抽出の手順", "品質管理", spec["bulletSlideTitle"], "3Dモデル配置"],
        2: ["医療事務研修", "受付の基本", "記録の取り扱い", "連絡体制", "個人情報", spec["bulletSlideTitle"], "3Dモデル配置"],
        3: ["設備保全研修", "保全の目的", "点検手順", "異常時対応", "記録管理", spec["bulletSlideTitle"], "3Dモデル配置"],
        4: ["旅行代理店研修", "接客の基本", "予約システム", "変更対応", "クレーム対応", spec["bulletSlideTitle"], "3Dモデル配置"],
        5: ["講師研修ガイド", "授業設計", "進度管理", "質問対応", "保護者連絡", spec["bulletSlideTitle"], "3Dモデル配置"],
    }[set_no]
    sm = []
    for i, t in enumerate(titles, 1):
        e = {"slideNo": i, "layout": "タイトルとコンテンツ", "title": t}
        if i == 2:
            e["notes"] = f"2-6: 画像「{spec['img1']}」「{spec['img2']}」配置・アニメ未設定"
            e["objects"] = [spec["img1"], spec["img2"]]
        if i == 4:
            e["notes"] = f"2-8: 図形「{spec['motionShape']}」に軌跡未設定・円図形はアニメなし"
            e["objects"] = [spec["motionShape"], "円"]
        if i == 6:
            e["notes"] = f"2-7: 箇条書きアニメ効果「{spec['bulletEffect']}」未変更"
            e["bullets"] = ["項目Aの確認", "項目Bの確認", "項目Cの確認"]
        if i == 7:
            e["objects"] = [spec["model3d"]]
            e["notes"] = f"2-5: 3Dモデル「{spec['model3d']}」に「{spec['modelAnim']}」未設定"
        sm.append(e)
    return sm


def _p2_tasks(spec: dict, intro: str) -> list[str]:
    t_all = spec["transitionAll"]
    opt = spec["transitionOption"]
    t1 = (
        f"タスク2-1　すべてのスライドに、画面切り替え「{t_all}」を設定します。"
        + (f"効果のオプションを「{opt}」にします。" if opt else "")
    )
    slides = "、".join(str(n) for n in spec["perSlideNos"])
    return [
        t1,
        f"タスク2-2　すべての画面切り替えの継続時間を{spec['durationSec']}秒に設定します。",
        f"タスク2-3　スライド{slides}に「{spec['perSlideEffect']}」の画面切り替え効果を設定します。",
        f"タスク2-4　すべてのスライドが、「{spec['autoAdvanceSec']}」秒後に自動で次のスライドへ進むように画面切り替えのタイミングを設定します。",
        f"タスク2-5　スライド６の3Ｄモデル「{spec['model3d']}」に、アニメーション「{spec['modelAnim']}」を設定します。",
        f"タスク2-6　スライド2の{spec['img1']}と{spec['img2']}の２つの画像が、スライドの{spec['animDir']}から登場するようにします。"
        f"アニメーションの継続時間は「{spec['animDur']}\"秒にします。",
        f"タスク2-7　スライド「{spec['bulletSlideTitle']}」の箇条書きに設定されたアニメーションの効果を「{spec['bulletEffect']}」に変更します。"
        "また、クリックするとすべて同時に動くようにします。",
        f"タスク2-8　スライド４の{spec['motionShape']}の図形にアニメーションの軌跡「{spec['motionPath']}」を設定します。円の図形にはアニメーションを設定しません。",
    ]


def _p2_task_params(spec: dict) -> dict:
    return {
        "task2_1": {
            "transition": spec["transitionAll"],
            "transitionOption": spec["transitionOption"] or None,
            "applyTo": "all",
            "operation": "setTransition",
        },
        "task2_2": {"durationSec": spec["durationSec"], "applyTo": "all", "operation": "transitionDuration"},
        "task2_3": {
            "slideNos": spec["perSlideNos"],
            "transition": spec["perSlideEffect"],
            "operation": "perSlideTransition",
        },
        "task2_4": {"autoAdvanceSec": spec["autoAdvanceSec"], "applyTo": "all", "operation": "autoAdvance"},
        "task2_5": {
            "slideNo": 6,
            "model3dName": spec["model3d"],
            "animation": spec["modelAnim"],
            "operation": "model3dAnimation",
        },
        "task2_6": {
            "slideNo": 2,
            "imageNames": [spec["img1"], spec["img2"]],
            "direction": spec["animDir"],
            "durationSec": spec["animDur"],
            "operation": "imageEntrance",
        },
        "task2_7": {
            "slideTitle": spec["bulletSlideTitle"],
            "slideNo": 6,
            "effect": spec["bulletEffect"],
            "timing": "すべて同時",
            "operation": "bulletAnimation",
        },
        "task2_8": {
            "slideNo": 4,
            "shapeName": spec["motionShape"],
            "excludeShape": "円",
            "motionPath": spec["motionPath"],
            "operation": "motionPath",
        },
    }


def build_project2_set(set_no: int, theme_meta: dict, spec: dict) -> dict:
    intro = P2_INTROS[set_no]
    return {
        "setNo": set_no,
        "theme": theme_meta["theme"],
        "workbook": f"MOS_PowerPoint類題_project2_配置別_セット{set_no}_{theme_meta['theme']}.pptx",
        "problemStatement": intro.split("\n")[0],
        "tasks": _p2_tasks(spec, intro),
        "layout": {
            "slideCount": 7,
            "slideSize": "ワイド スクリーン (16:9)",
            "designTheme": theme_meta["designTheme"],
            "variant": theme_meta["variant"],
            "colorTheme": theme_meta["colorTheme"],
            "fontTheme": theme_meta["fontTheme"],
        },
        "contentBlocks": {
            "slideMap": _p2_slide_map(spec, set_no),
            "preTaskState": {
                "transitionsApplied": False,
                "animationsApplied": False,
                "autoAdvanceSet": False,
            },
        },
        "taskParams": _p2_task_params(spec),
    }


# ---------------------------------------------------------------------------
# Project 3
# ---------------------------------------------------------------------------

P3_INTROS = {
    1: "あなたは店舗DX導入についての社内説明資料を作成しています。",
    2: "あなたは電子カルテ運用についての院内説明資料を作成しています。",
    3: "あなたは生産ラインIoT化についての説明資料を作成しています。",
    4: "あなたは予約システム刷新についての説明資料を作成しています。",
    5: "あなたは学習管理システム導入についての説明資料を作成しています。",
}

P3_SETS = [
    {
        "smartArtType": "基本プロセス", "smartArtLabels": ["受付", "提供"],
        "smartArtColor": "塗りつぶし-アクセント3",
        "convertSmartArt": "水平ボックスリスト",
        "modelInsert": "ノートパソコン", "modelWidth": 3.0,
        "modelView": "上面", "modelHeight": 5.5,
        "zoomParentTitle": "導入の概要",
        "zoomLinks": ["POS連携の手順", "スタッフ教育の流れ", "売上分析の見える化"],
        "sectionZoom": ["1.導入の概要", "2.運用のポイント"],
        "sec1": "1.導入の概要", "sec2": "2.運用のポイント",
    },
    {
        "smartArtType": "上向き矢印", "smartArtLabels": ["問診", "診察"],
        "smartArtColor": "カラフル-アクセント4",
        "convertSmartArt": "箇条書き",
        "modelInsert": "聴診器", "modelWidth": 2.8,
        "modelView": "上前面", "modelHeight": 5.0,
        "zoomParentTitle": "運用の全体像",
        "zoomLinks": ["記録入力の手順", "権限設定の流れ", "監査ログの確認"],
        "sectionZoom": ["1.運用の全体像", "2.セキュリティ対策"],
        "sec1": "1.運用の全体像", "sec2": "2.セキュリティ対策",
    },
    {
        "smartArtType": "連続ブロック プロセス", "smartArtLabels": ["収集", "分析"],
        "smartArtColor": "塗りつぶし-アクセント2",
        "convertSmartArt": "L型リスト",
        "modelInsert": "工具箱", "modelWidth": 3.2,
        "modelView": "上面", "modelHeight": 6.0,
        "zoomParentTitle": "IoT導入の概要",
        "zoomLinks": ["センサー設置手順", "データ連携手順", "異常検知の運用"],
        "sectionZoom": ["1.IoT導入の概要", "2.保全の効率化"],
        "sec1": "1.IoT導入の概要", "sec2": "2.保全の効率化",
    },
    {
        "smartArtType": "基本プロセス", "smartArtLabels": ["予約", "決済"],
        "smartArtColor": "グラデーション-アクセント5",
        "convertSmartArt": "横方向のリスト",
        "modelInsert": "飛行機", "modelWidth": 2.7,
        "modelView": "上前面", "modelHeight": 5.8,
        "zoomParentTitle": "刷新計画の概要",
        "zoomLinks": ["在庫連携の手順", "顧客通知の流れ", "レポート活用"],
        "sectionZoom": ["1.刷新計画の概要", "2.オペレーション改善"],
        "sec1": "1.刷新計画の概要", "sec2": "2.オペレーション改善",
    },
    {
        "smartArtType": "上向き矢印", "smartArtLabels": ["登録", "共有"],
        "smartArtColor": "カラフル-アクセント1",
        "convertSmartArt": "水平ボックスリスト",
        "modelInsert": "ノートパソコン", "modelWidth": 3.1,
        "modelView": "上面", "modelHeight": 5.2,
        "zoomParentTitle": "LMS導入の概要",
        "zoomLinks": ["教材配布の手順", "課題提出の流れ", "成績集計の活用"],
        "sectionZoom": ["1.LMS導入の概要", "2.学習データの保護"],
        "sec1": "1.LMS導入の概要", "sec2": "2.学習データの保護",
    },
]


def _p3_slide_map(spec: dict, set_no: int) -> list[dict]:
    base_titles = [
        (1, "タイトルスライド", f"{P3_INTROS[set_no].replace('あなたは', '').replace('についての説明資料を作成しています。', '')}"),
        (2, "タイトルとコンテンツ", "目次"),
        (3, "タイトルとコンテンツ", spec["zoomParentTitle"]),
        (4, "タイトルとコンテンツ", spec["zoomLinks"][0]),
        (5, "タイトルとコンテンツ", spec["zoomLinks"][1]),
        (6, "タイトルとコンテンツ", spec["zoomLinks"][2]),
        (7, "タイトルとコンテンツ", "SmartArt配置スライド"),
        (8, "タイトルとコンテンツ", spec["sec1"]),
        (9, "タイトルとコンテンツ", "補足"),
        (10, "タイトルとコンテンツ", "3Dモデル配置スライド"),
    ]
    sm = []
    for no, layout, title in base_titles:
        e = {"slideNo": no, "layout": layout, "title": title}
        if no == 1:
            e["notes"] = "3-4: 楕円図形内に3Dモデル未挿入"
            e["objects"] = ["楕円図形"]
        if no == 3:
            e["notes"] = "3-6: スライドズーム未挿入"
        if no == 2:
            e["notes"] = "3-7: セクションズーム未挿入"
        if no == 6:
            e["bullets"] = ["要点A", "要点B", "要点C"]
            e["notes"] = f"3-3: 箇条書き→SmartArt「{spec['convertSmartArt']}」未変換"
        if no == 7:
            e["notes"] = f"3-1/3-2: SmartArt「{spec['smartArtType']}」未追加・色未設定"
        if no == 10:
            e["objects"] = ["既存3Dモデル"]
            e["notes"] = "3-5: ビュー・高さ未変更"
        sm.append(e)
    return sm


def _p3_tasks(spec: dict, intro: str) -> list[str]:
    la, lb = spec["smartArtLabels"]
    zl = spec["zoomLinks"]
    return [
        f"タスク3-1　スライド７にSmartArtグラフィック「{spec['smartArtType']}」の手順を追加し、"
        f"文字列「\"{la}\"」「\"{lb}\"」を入力します。不要な図形は削除します。",
        f"タスク3-2　スライド７のSmartArtグラフィックに、色「{spec['smartArtColor']}」を設定します。",
        f"タスク3-3　スライド６の箇条書きを「{spec['convertSmartArt']}」のSmartArtに変更します。",
        f"タスク3-4　スライド１に3Dモデル「{spec['modelInsert']}」を挿入します。幅を「\"{spec['modelWidth']}\"」に変更して"
        "中央の楕円の図形の中に配置します。正確な位置は問いません。",
        f"タスク3-5　スライド10の3Dモデルのビューを{spec['modelView']}にし、高さを「\"{spec['modelHeight']}\"」に変更します。",
        f"タスク3-6　「{spec['zoomParentTitle']}」のスライドにスライドズームを挿入して"
        f"「{zl[0]}」「{zl[1]}」「{zl[2]}」へリンクを作成します。タイトルの下に配置し、スライドズームが重ならないようにします。",
        f"タスク3-7　スライド２にセクションズームのリンクを挿入します。セクション「{spec['sectionZoom'][0]}」と"
        f"「{spec['sectionZoom'][1]}」にリンクを作成し、それぞれ文字の下に配置します。",
    ]


def _p3_task_params(spec: dict) -> dict:
    return {
        "task3_1": {
            "slideNo": 7,
            "smartArtCategory": "手順",
            "smartArtType": spec["smartArtType"],
            "labels": spec["smartArtLabels"],
            "operation": "insertSmartArt",
        },
        "task3_2": {
            "slideNo": 7,
            "smartArtColor": spec["smartArtColor"],
            "operation": "smartArtColor",
        },
        "task3_3": {
            "slideNo": 6,
            "smartArtType": spec["convertSmartArt"],
            "smartArtCategory": "リスト",
            "operation": "convertToSmartArt",
        },
        "task3_4": {
            "slideNo": 1,
            "modelName": spec["modelInsert"],
            "width": spec["modelWidth"],
            "anchorShape": "楕円",
            "operation": "insert3dModel",
        },
        "task3_5": {
            "slideNo": 10,
            "view": spec["modelView"],
            "height": spec["modelHeight"],
            "operation": "model3dViewSize",
        },
        "task3_6": {
            "parentSlideTitle": spec["zoomParentTitle"],
            "linkTitles": spec["zoomLinks"],
            "operation": "slideZoom",
        },
        "task3_7": {
            "slideNo": 2,
            "sectionNames": spec["sectionZoom"],
            "operation": "sectionZoom",
        },
    }


def build_project3_set(set_no: int, theme_meta: dict, spec: dict) -> dict:
    intro = P3_INTROS[set_no]
    return {
        "setNo": set_no,
        "theme": theme_meta["theme"],
        "workbook": f"MOS_PowerPoint類題_project3_配置別_セット{set_no}_{theme_meta['theme']}.pptx",
        "problemStatement": intro.split("\n")[0],
        "tasks": _p3_tasks(spec, intro),
        "layout": {
            "slideCount": 10,
            "slideSize": "ワイド スクリーン (16:9)",
            "designTheme": theme_meta["designTheme"],
            "variant": theme_meta["variant"],
            "colorTheme": theme_meta["colorTheme"],
            "fontTheme": theme_meta["fontTheme"],
            "sections": [
                {"startSlide": 1, "name": spec["sec1"]},
                {"startSlide": 8, "name": spec["sec2"]},
            ],
        },
        "contentBlocks": {
            "slideMap": _p3_slide_map(spec, set_no),
            "preTaskState": {
                "smartArtOnSlide7": False,
                "slideZoomOnSlide3": False,
                "sectionZoomOnSlide2": False,
            },
        },
        "taskParams": _p3_task_params(spec),
    }


# ---------------------------------------------------------------------------
# Project 4
# ---------------------------------------------------------------------------

P4_INTROS = {
    1: "あなたは店舗メニュー紹介資料を作成しています。",
    2: "あなたは健診案内チラシ資料を作成しています。",
    3: "あなたは工場見学会案内資料を作成しています。",
    4: "あなたはツアー商品紹介資料を作成しています。",
    5: "あなたは講座案内チラシ資料を作成しています。",
}

P4_SETS = [
    {
        "inputShapeType": "角丸四角", "inputShapeInsertName": "角丸四角形",
        "inputShapeColorDesc": "茶", "shapeText": "スタッフ必読",
        "targetPhrase": "いつでも試飲できます!", "fillColor": "オレンジ、アクセント4",
        "imageLabel": "店舗外観", "glowPt": 12, "glowColor": "オレンジ、アクセントカラー4",
        "quickStyle": "透かし", "artEffect": "ブロック",
        "imageLeft": "メニュー看板", "imageRight": "カップ写真",
        "textBoxShapeType": "角丸四角形", "textBoxInsertName": "角丸四角形",
        "shapeFillColor": "茶、アクセント1、白+基本色60%", "borderColor": "濃い茶", "borderPt": 1,
        "slideTitles": ["春のメニュー案内", "提供スタイル", "店舗コンセプト", "体験キャンペーン", "商品写真"],
    },
    {
        "inputShapeType": "平行四辺形", "inputShapeInsertName": "平行四辺形",
        "inputShapeColorDesc": "緑", "shapeText": "受付案内",
        "targetPhrase": "平日夜間も受付中!", "fillColor": "緑、アクセント1",
        "imageLabel": "受付写真", "glowPt": 15, "glowColor": "青、アクセントカラー5",
        "quickStyle": "反射", "artEffect": "水彩",
        "imageLeft": "待合室", "imageRight": "案内板",
        "textBoxShapeType": "テキストボックス", "textBoxInsertName": "テキストボックス",
        "shapeFillColor": "緑、アクセント1、白+基本色60%", "borderColor": "濃い緑", "borderPt": 1.5,
        "slideTitles": ["健診メニュー案内", "受付の流れ", "検査の準備", "夜間受付案内", "院内写真"],
    },
    {
        "inputShapeType": "六角形", "inputShapeInsertName": "六角形",
        "inputShapeColorDesc": "青灰", "shapeText": "見学歓迎",
        "targetPhrase": "安全装備をご用意ください!", "fillColor": "青灰、アクセント2",
        "imageLabel": "工場全景", "glowPt": 20, "glowColor": "紫、アクセントカラー3",
        "quickStyle": "ソフトエッジ 長方形", "artEffect": "フィルムの粒",
        "imageLeft": "ライン設備", "imageRight": "安全標識",
        "textBoxShapeType": "平行四辺形", "textBoxInsertName": "平行四辺形",
        "shapeFillColor": "青灰、アクセント2、白+基本色60%", "borderColor": "濃い青", "borderPt": 2,
        "slideTitles": ["見学会のご案内", "見学ルート", "安全の心得", "持ち物のご案内", "工場写真"],
    },
    {
        "inputShapeType": "クラウド", "inputShapeInsertName": "クラウド",
        "inputShapeColorDesc": "水色", "shapeText": "予約受付中",
        "targetPhrase": "早期割引あり!", "fillColor": "水色、アクセント5",
        "imageLabel": "旅行風景", "glowPt": 22, "glowColor": "赤、アクセントカラー2",
        "quickStyle": "角丸四角形", "artEffect": "マーカー",
        "imageLeft": "観光地", "imageRight": "旅行カバン",
        "textBoxShapeType": "五角形", "textBoxInsertName": "五角形",
        "shapeFillColor": "水色、アクセント5、白+基本色60%", "borderColor": "濃い水色", "borderPt": 2.5,
        "slideTitles": ["春のツアー案内", "行程の概要", "持ち物リスト", "割引キャンペーン", "旅行写真"],
    },
    {
        "inputShapeType": "吹き出し: 角丸", "inputShapeInsertName": "吹き出し: 角丸",
        "inputShapeColorDesc": "オレンジ", "shapeText": "体験申込可",
        "targetPhrase": "無料体験実施中!", "fillColor": "紫、アクセント3",
        "imageLabel": "教室風景", "glowPt": 25, "glowColor": "黄色",
        "quickStyle": "ぼかし 長方形", "artEffect": "ペイントストローク",
        "imageLeft": "講義風景", "imageRight": "教材サンプル",
        "textBoxShapeType": "六角形", "textBoxInsertName": "六角形",
        "shapeFillColor": "オレンジ、アクセント4、白+基本色60%", "borderColor": "濃いオレンジ", "borderPt": 3,
        "slideTitles": ["秋期講座案内", "授業の流れ", "学習方針", "体験授業案内", "教室写真"],
    },
]


def _p4_slide_map(spec: dict, set_no: int) -> list[dict]:
    titles = spec["slideTitles"]
    sm = []
    for i, title in enumerate(titles, 1):
        e: dict = {"slideNo": i, "layout": "タイトルとコンテンツ", "title": title}
        if i == 1:
            e["shapes"] = [
                {
                    "id": "inputShape",
                    "type": spec["inputShapeType"],
                    "insertName": spec["inputShapeInsertName"],
                    "fillColorDesc": spec["inputShapeColorDesc"],
                    "count": 1,
                    "text": None,
                    "notes": "4-1: 文字未入力",
                },
                {
                    "id": "heroImage",
                    "type": "画像",
                    "insertName": "画像",
                    "label": spec["imageLabel"],
                    "count": 1,
                    "notes": "4-3/4-4: 光彩・スタイル・アート効果未設定",
                },
            ]
            e["objects"] = [spec["inputShapeType"], spec["imageLabel"]]
            e["notes"] = "4-1: 図形に文字未入力 / 4-3/4-4: 画像効果未設定"
        elif i == 2:
            e["shapes"] = [
                {
                    "id": "textBoxShape",
                    "type": spec["textBoxShapeType"],
                    "insertName": spec["textBoxInsertName"],
                    "count": 1,
                    "notes": "4-7: 塗り・枠線未設定",
                },
            ]
            e["notes"] = "4-7: テキストボックスの書式未設定"
        elif i == 3:
            e["shapes"] = [
                {
                    "id": "centerTextBox",
                    "type": spec["textBoxShapeType"],
                    "insertName": spec["textBoxInsertName"],
                    "count": 1,
                    "notes": "4-8: 上下中央配置未実施",
                },
            ]
            e["notes"] = "4-8: 垂直方向中央配置未実施"
        elif i == 4:
            e["shapes"] = [
                {
                    "id": "phraseText",
                    "type": "テキストボックス",
                    "insertName": "テキストボックス",
                    "text": spec["targetPhrase"],
                    "count": 1,
                    "notes": "4-2: 文字塗りつぶし未設定",
                },
            ]
            e["notes"] = f"4-2: 「{spec['targetPhrase']}」の塗りつぶし未設定"
        elif i == 5:
            e["shapes"] = [
                {"id": "imageLeft", "type": "画像", "insertName": "画像", "label": spec["imageLeft"], "count": 1},
                {"id": "imageRight", "type": "画像", "insertName": "画像", "label": spec["imageRight"], "count": 1},
            ]
            e["objects"] = [spec["imageLeft"], spec["imageRight"]]
            e["notes"] = "4-5/4-6: 画像配置・トリミング未実施"
        sm.append(e)
    return sm


def _p4_tasks(spec: dict) -> list[str]:
    glow = f"光彩:{spec['glowPt']}pt;{spec['glowColor']}"
    return [
        f"タスク4-1　スライド１の{spec['inputShapeColorDesc']}の{spec['inputShapeType']}に、"
        f"「\"{spec['shapeText']}\"」と入力します。",
        f"タスク4-2　スライド４の文字列「{spec['targetPhrase']}」に、文字の塗りつぶし「{spec['fillColor']}」を設定します。",
        f"タスク4-3　スライド１枚目の{spec['imageLabel']}の画像に、図の効果「{glow}」を設定します。",
        f"タスク4-4　スライド１の{spec['imageLabel']}の画像に、スタイル「{spec['quickStyle']}」を設定し、"
        f"「{spec['artEffect']}」のアート効果を設定します。",
        "タスク4-5　スライド５の右の画像を左の画像の上端に合わせます。水平方向には動かさないようにします。",
        "タスク4-6　スライド５の右側の画像の右端を、スライドの右端に揃えてトリミングします。"
        "画像の右端以外は変更しないでください。トリミングした領域をプレゼンテーションから完全に削除しないでください。",
        f"タスク4-7　2枚目のスライドの{spec['textBoxShapeType']}に、塗りつぶし「{spec['shapeFillColor']}」、"
        f"枠線「{spec['borderColor']}」、太さ「{spec['borderPt']}pt」を設定します。",
        "タスク4-8　スライド３のコンテンツ領域にあるテキストボックスを、スライドの垂直方向の中央に配置します。",
    ]


def _p4_task_params(spec: dict) -> dict:
    return {
        "task4_1": {
            "slideNo": 1,
            "inputShapeType": spec["inputShapeType"],
            "inputShapeInsertName": spec["inputShapeInsertName"],
            "inputShapeColorDesc": spec["inputShapeColorDesc"],
            "shapeText": spec["shapeText"],
            "operation": "shapeTextInput",
        },
        "task4_2": {
            "slideNo": 4,
            "targetPhrase": spec["targetPhrase"],
            "fillColor": spec["fillColor"],
            "operation": "textFillColor",
        },
        "task4_3": {
            "slideNo": 1,
            "imageLabel": spec["imageLabel"],
            "glowPt": spec["glowPt"],
            "glowColor": spec["glowColor"],
            "operation": "pictureGlow",
        },
        "task4_4": {
            "slideNo": 1,
            "imageLabel": spec["imageLabel"],
            "quickStyle": spec["quickStyle"],
            "artEffect": spec["artEffect"],
            "operation": "pictureStyleArt",
        },
        "task4_5": {"slideNo": 5, "align": "上揃え", "operation": "alignImages"},
        "task4_6": {"slideNo": 5, "side": "right", "operation": "cropImage"},
        "task4_7": {
            "slideNo": 2,
            "textBoxShapeType": spec["textBoxShapeType"],
            "shapeFillColor": spec["shapeFillColor"],
            "borderColor": spec["borderColor"],
            "borderPt": spec["borderPt"],
            "operation": "shapeFillBorder",
        },
        "task4_8": {"slideNo": 3, "align": "上下中央揃え", "operation": "verticalCenter"},
    }


def build_project4_set(set_no: int, theme_meta: dict, spec: dict) -> dict:
    intro = P4_INTROS[set_no]
    return {
        "setNo": set_no,
        "theme": theme_meta["theme"],
        "workbook": f"MOS_PowerPoint類題_project4_配置別_セット{set_no}_{theme_meta['theme']}.pptx",
        "problemStatement": intro,
        "tasks": _p4_tasks(spec),
        "layout": {
            "slideCount": 5,
            "slideSize": "ワイド スクリーン (16:9)",
            "designTheme": theme_meta["designTheme"],
            "variant": theme_meta["variant"],
            "colorTheme": theme_meta["colorTheme"],
            "fontTheme": theme_meta["fontTheme"],
        },
        "contentBlocks": {
            "slideMap": _p4_slide_map(spec, set_no),
            "preTaskState": {
                "shapeTextOnSlide1": False,
                "textFillOnSlide4": False,
                "pictureEffectsOnSlide1": False,
                "imageAlignOnSlide5": False,
                "textBoxFormatOnSlide2": False,
                "verticalCenterOnSlide3": False,
            },
        },
        "taskParams": _p4_task_params(spec),
    }


# ---------------------------------------------------------------------------
# Project 5
# ---------------------------------------------------------------------------

P5_INTROS = {
    1: "あなたはバリスタ研修資料を作成しています。",
    2: "あなたは医療事務研修資料を作成しています。",
    3: "あなたは設備保全研修資料を作成しています。",
    4: "あなたは旅行代理店研修資料を作成しています。",
    5: "あなたは講師研修資料を作成しています。",
}

P5_SETS = [
    {
        "alignShapeType": "六角形", "alignShapeInsertName": "六角形",
        "rectShapeType": "五角形", "rectShapeInsertName": "五角形",
        "sourceShapeType": "雲", "sourceShapeInsertName": "雲",
        "targetShapeType": "ハート", "targetShapeInsertName": "ハート",
        "zOrderShapeType": "四角形", "zOrderShapeInsertName": "四角形",
        "zOrderLabels": ["バリスタ研修", "抽出実習", "品質チェック"],
        "gateShapeType": "論理積ゲート", "gateShapeInsertName": "論理積ゲート",
        "decorativeSlideTitle": "研修の目的", "decorativeImageLabel": "講師写真",
        "iconSymbol": "?", "iconFillColor": "濃い青",
        "slideTitles": ["バリスタ研修ガイド", "研修の目的", "工程の確認", "品質管理", "実習項目", "研修フロー"],
    },
    {
        "alignShapeType": "四角", "alignShapeInsertName": "四角形",
        "rectShapeType": "角丸四角", "rectShapeInsertName": "角丸四角形",
        "sourceShapeType": "ハート", "sourceShapeInsertName": "ハート",
        "targetShapeType": "太陽", "targetShapeInsertName": "太陽",
        "zOrderShapeType": "角丸四角", "zOrderShapeInsertName": "角丸四角形",
        "zOrderLabels": ["初診受付", "検査予約", "結果説明"],
        "gateShapeType": "論理和ゲート", "gateShapeInsertName": "論理和ゲート",
        "decorativeSlideTitle": "受付の基本", "decorativeImageLabel": "スタッフ写真",
        "iconSymbol": "!", "iconFillColor": "オレンジ",
        "slideTitles": ["医療事務研修", "受付の基本", "記録管理", "連絡体制", "業務分担", "受付フロー"],
    },
    {
        "alignShapeType": "三角", "alignShapeInsertName": "三角形",
        "rectShapeType": "平行四辺形", "rectShapeInsertName": "平行四辺形",
        "sourceShapeType": "稲妻", "sourceShapeInsertName": "稲妻",
        "targetShapeType": "月", "targetShapeInsertName": "月",
        "zOrderShapeType": "吹き出し", "zOrderShapeInsertName": "吹き出し: 角丸",
        "zOrderLabels": ["日常点検", "定期保全", "異常対応"],
        "gateShapeType": "排他的論理和", "gateShapeInsertName": "排他的論理和",
        "decorativeSlideTitle": "保全の目的", "decorativeImageLabel": "技術者写真",
        "iconSymbol": "i", "iconFillColor": "紫",
        "slideTitles": ["設備保全研修", "保全の目的", "保全手順", "異常対応", "記録管理", "保全フロー"],
    },
    {
        "alignShapeType": "菱形", "alignShapeInsertName": "菱形",
        "rectShapeType": "ひし形", "rectShapeInsertName": "菱形",
        "sourceShapeType": "月", "sourceShapeInsertName": "月",
        "targetShapeType": "花", "targetShapeInsertName": "花",
        "zOrderShapeType": "四角形", "zOrderShapeInsertName": "四角形",
        "zOrderLabels": ["国内旅行", "海外旅行", "添乗員同行"],
        "gateShapeType": "論理和ゲート", "gateShapeInsertName": "論理和ゲート",
        "decorativeSlideTitle": "接客の基本", "decorativeImageLabel": "カウンター写真",
        "iconSymbol": "+", "iconFillColor": "黄色",
        "slideTitles": ["旅行代理店研修", "接客の基本", "予約対応", "変更手続", "クレーム対応", "予約フロー"],
    },
    {
        "alignShapeType": "五角形", "alignShapeInsertName": "五角形",
        "rectShapeType": "六角形", "rectShapeInsertName": "六角形",
        "sourceShapeType": "平行四辺形", "sourceShapeInsertName": "平行四辺形",
        "targetShapeType": "稲妻", "targetShapeInsertName": "稲妻",
        "zOrderShapeType": "角丸四角", "zOrderShapeInsertName": "角丸四角形",
        "zOrderLabels": ["春期講座", "夏期講座", "秋期講座"],
        "gateShapeType": "論理和ゲート", "gateShapeInsertName": "論理和ゲート",
        "decorativeSlideTitle": "授業設計", "decorativeImageLabel": "講師写真",
        "iconSymbol": "※", "iconFillColor": "濃い緑",
        "slideTitles": ["講師研修ガイド", "授業設計", "進度管理", "質問対応", "保護者連絡", "授業フロー"],
    },
]


def _p5_slide_map(spec: dict, set_no: int) -> list[dict]:
    titles = spec["slideTitles"]
    sm = []
    for i, title in enumerate(titles, 1):
        e: dict = {"slideNo": i, "layout": "タイトルとコンテンツ", "title": title}
        if i == 2:
            e["shapes"] = [
                {
                    "id": "decorativeImage",
                    "type": "画像",
                    "insertName": "画像",
                    "label": spec["decorativeImageLabel"],
                    "count": 1,
                    "notes": "5-6: 代替テキスト装飾化未実施",
                },
                {
                    "id": "helpIcon",
                    "type": "アイコン",
                    "insertName": "アイコン",
                    "label": f"[{spec['iconSymbol']}]",
                    "count": 1,
                    "notes": f"5-7: 塗りつぶし「{spec['iconFillColor']}」未設定",
                },
            ]
            e["notes"] = "5-6/5-7: 代替テキスト・アイコン塗り未設定"
        elif i == 3:
            e["shapes"] = [
                {
                    "id": f"alignShape{n}",
                    "type": spec["alignShapeType"],
                    "insertName": spec["alignShapeInsertName"],
                    "count": 1,
                }
                for n in range(1, 5)
            ]
            e["notes"] = "5-1: 4図形の右揃え未実施"
        elif i == 4:
            e["shapes"] = [
                {
                    "id": "sourceShape",
                    "type": spec["sourceShapeType"],
                    "insertName": spec["sourceShapeInsertName"],
                    "count": 1,
                    "notes": f"5-3: 「{spec['targetShapeType']}」への変更未実施",
                },
            ]
            e["notes"] = f"5-3: {spec['sourceShapeType']}→{spec['targetShapeType']} 未変更"
        elif i == 5:
            e["shapes"] = [
                {
                    "id": "rectLarge1",
                    "type": spec["rectShapeType"],
                    "insertName": spec["rectShapeInsertName"],
                    "count": 1,
                    "label": "大",
                },
                {
                    "id": "rectLarge2",
                    "type": spec["rectShapeType"],
                    "insertName": spec["rectShapeInsertName"],
                    "count": 1,
                    "label": "大",
                },
                {
                    "id": "rectSmall",
                    "type": spec["rectShapeType"],
                    "insertName": spec["rectShapeInsertName"],
                    "count": 1,
                    "label": "小",
                    "notes": "5-2: 幅が他と異なる",
                },
            ]
            e["notes"] = "5-2: 小さい図形の幅調整未実施"
        elif i == 6:
            z_shapes = []
            for n, label in enumerate(spec["zOrderLabels"], 1):
                z_shapes.append({
                    "id": f"zOrderShape{n}",
                    "type": spec["zOrderShapeType"],
                    "insertName": spec["zOrderShapeInsertName"],
                    "label": label,
                    "count": 1,
                    "notes": "5-4: 重なり順未変更",
                })
            for n in range(1, 4):
                z_shapes.append({
                    "id": f"gateShape{n}",
                    "type": spec["gateShapeType"],
                    "insertName": spec["gateShapeInsertName"],
                    "count": 1,
                    "notes": "5-5: グループ化未実施",
                })
            e["shapes"] = z_shapes
            e["notes"] = "5-4/5-5: 重なり順・グループ化未実施"
        sm.append(e)
    return sm


def _p5_tasks(spec: dict) -> list[str]:
    l1, l2, l3 = spec["zOrderLabels"]
    return [
        f"タスク5-1　スライド３の{spec['alignShapeType']}の図形４個の右端を揃えます。",
        f"タスク5-2　スライド５の小さい{spec['rectShapeType']}の大きさを他の{spec['rectShapeType']}と同じ幅にします。",
        f"タスク5-3　スライド４の{spec['sourceShapeType']}の図形を{spec['targetShapeType']}に変更します。",
        f"タスク5-4　スライド６の図形を手前から「{l1}」「{l2}」「{l3}」になるように重なりの順番を変更します。",
        f"タスク5-5　スライド６の3つの{spec['gateShapeType']}の図形をグループ化します。",
        f"タスク5-6　スライド「{spec['decorativeSlideTitle']}」の{spec['decorativeImageLabel']}の画像の代替テキストを装飾化し、"
        "スクリーンリーダーに表示させないようにします。",
        f"タスク5-7　スライド２の[{spec['iconSymbol']}]のアイコンに「{spec['iconFillColor']}」の塗りつぶしを設定します。",
    ]


def _p5_task_params(spec: dict) -> dict:
    return {
        "task5_1": {
            "slideNo": 3,
            "alignShapeType": spec["alignShapeType"],
            "alignShapeInsertName": spec["alignShapeInsertName"],
            "count": 4,
            "operation": "alignRight",
        },
        "task5_2": {
            "slideNo": 5,
            "rectShapeType": spec["rectShapeType"],
            "rectShapeInsertName": spec["rectShapeInsertName"],
            "operation": "matchWidth",
        },
        "task5_3": {
            "slideNo": 4,
            "sourceShapeType": spec["sourceShapeType"],
            "sourceShapeInsertName": spec["sourceShapeInsertName"],
            "targetShapeType": spec["targetShapeType"],
            "targetShapeInsertName": spec["targetShapeInsertName"],
            "operation": "changeShape",
        },
        "task5_4": {
            "slideNo": 6,
            "zOrderShapeType": spec["zOrderShapeType"],
            "labels": spec["zOrderLabels"],
            "operation": "zOrder",
        },
        "task5_5": {
            "slideNo": 6,
            "gateShapeType": spec["gateShapeType"],
            "gateShapeInsertName": spec["gateShapeInsertName"],
            "count": 3,
            "operation": "groupShapes",
        },
        "task5_6": {
            "slideTitle": spec["decorativeSlideTitle"],
            "imageLabel": spec["decorativeImageLabel"],
            "operation": "decorativeAltText",
        },
        "task5_7": {
            "slideNo": 2,
            "iconSymbol": spec["iconSymbol"],
            "iconFillColor": spec["iconFillColor"],
            "operation": "iconFill",
        },
    }


def build_project5_set(set_no: int, theme_meta: dict, spec: dict) -> dict:
    intro = P5_INTROS[set_no]
    return {
        "setNo": set_no,
        "theme": theme_meta["theme"],
        "workbook": f"MOS_PowerPoint類題_project5_配置別_セット{set_no}_{theme_meta['theme']}.pptx",
        "problemStatement": intro,
        "tasks": _p5_tasks(spec),
        "layout": {
            "slideCount": 6,
            "slideSize": "ワイド スクリーン (16:9)",
            "designTheme": theme_meta["designTheme"],
            "variant": theme_meta["variant"],
            "colorTheme": theme_meta["colorTheme"],
            "fontTheme": theme_meta["fontTheme"],
        },
        "contentBlocks": {
            "slideMap": _p5_slide_map(spec, set_no),
            "preTaskState": {
                "alignOnSlide3": False,
                "widthMatchOnSlide5": False,
                "shapeChangeOnSlide4": False,
                "zOrderOnSlide6": False,
                "groupOnSlide6": False,
                "decorativeAltText": False,
                "iconFillOnSlide2": False,
            },
        },
        "taskParams": _p5_task_params(spec),
    }


# ---------------------------------------------------------------------------
# Project 6
# ---------------------------------------------------------------------------

P6_INTROS = {
    1: "あなたは接客プレゼン技法資料を作成しています。",
    2: "あなたは院内説明会資料を作成しています。",
    3: "あなたは安全啓発セミナー資料を作成しています。",
    4: "あなたは提案力向上セミナー資料を作成しています。",
    5: "あなたは授業運営セミナー資料を作成しています。",
}

P6_SETS = [
    {
        "customShowName": "デザインの要点",
        "outlineCopies": 4, "notesCopies": 2, "handoutCopies": 2,
        "slideTitles": [
            "接客プレゼンの基本", "声の出し方", "視線の配り方", "デザインの考え方",
            "配色のポイント", "レイアウトの工夫", "まとめ",
        ],
    },
    {
        "customShowName": "構成の要点",
        "outlineCopies": 5, "notesCopies": 4, "handoutCopies": 3,
        "slideTitles": [
            "院内説明会の進め方", "資料の準備", "質疑応答", "構成の考え方",
            "情報整理のコツ", "伝え方の工夫", "まとめ",
        ],
    },
    {
        "customShowName": "安全の要点",
        "outlineCopies": 7, "notesCopies": 5, "handoutCopies": 5,
        "slideTitles": [
            "安全啓発の基本", "危険予知", "保護具の着用", "安全確認の手順",
            "事故防止策", "報告の流れ", "まとめ",
        ],
    },
    {
        "customShowName": "提案の要点",
        "outlineCopies": 8, "notesCopies": 6, "handoutCopies": 6,
        "slideTitles": [
            "提案力向上の基本", "顧客分析", "商品選定", "提案の流れ",
            "訴求ポイント", "クロージング", "まとめ",
        ],
    },
    {
        "customShowName": "運営の要点",
        "outlineCopies": 10, "notesCopies": 8, "handoutCopies": 8,
        "slideTitles": [
            "授業運営の基本", "時間配分", "教材準備", "運営の手順",
            "進度管理", "振り返り", "まとめ",
        ],
    },
]


def _p6_slide_map(spec: dict) -> list[dict]:
    sm = []
    for i, title in enumerate(spec["slideTitles"], 1):
        e: dict = {"slideNo": i, "layout": "タイトルとコンテンツ", "title": title}
        if i in (4, 5, 6):
            e["notes"] = f"6-4: 目的別スライドショー「{spec['customShowName']}」の対象"
        sm.append(e)
    return sm


def _p6_tasks(spec: dict) -> list[str]:
    return [
        "タスク6-1　プレゼンテーションからドキュメントのプロパティと個人情報を削除します。"
        "他の情報は削除しないでください。「ドキュメント検査の前にファイルを保存します。」のメッセージが表示された場合は「はい」をクリックします。",
        "タスク6-2　プレゼンテーションを常に読み取り専用にします。",
        "タスク6-3　スライドショーを自動プレゼンテーションとして設定します。",
        f"タスク6-4　スライド４，５，６を選択し、「\"{spec['customShowName']}\"」という名前の目的別スライドショーを作成します。スライドショーは実行しません。",
        f"タスク6-5　すべてのスライドをアウトラインで、部単位に\"{spec['outlineCopies']}\"部印刷するように設定します。",
        f"タスク6-6　ノートですべてのスライドを\"{spec['notesCopies']}\"部印刷しなさい。"
        "ただし、1ページ目を全て印刷したあとに2ページ目を印刷するようにします。",
        f"タスク6-7　印刷オプションで、グレースケールの配布資料を、１ページに3スライドのレイアウトで"
        f"\"{spec['handoutCopies']}\"部印刷するように設定します。印刷は実行しないでください。",
    ]


def _p6_task_params(spec: dict) -> dict:
    return {
        "task6_1": {"operation": "documentInspect", "remove": ["properties", "personalInfo"]},
        "task6_2": {"operation": "readOnly"},
        "task6_3": {"operation": "autoPresentation"},
        "task6_4": {
            "slideNos": [4, 5, 6],
            "customShowName": spec["customShowName"],
            "operation": "customSlideShow",
        },
        "task6_5": {"outlineCopies": spec["outlineCopies"], "collate": "部単位", "operation": "printOutline"},
        "task6_6": {"notesCopies": spec["notesCopies"], "collate": "ページ単位", "operation": "printNotes"},
        "task6_7": {
            "handoutCopies": spec["handoutCopies"],
            "slidesPerPage": 3,
            "grayscale": True,
            "operation": "printHandout",
        },
    }


def build_project6_set(set_no: int, theme_meta: dict, spec: dict) -> dict:
    intro = P6_INTROS[set_no]
    return {
        "setNo": set_no,
        "theme": theme_meta["theme"],
        "workbook": f"MOS_PowerPoint類題_project6_配置別_セット{set_no}_{theme_meta['theme']}.pptx",
        "problemStatement": intro,
        "tasks": _p6_tasks(spec),
        "layout": {
            "slideCount": len(spec["slideTitles"]),
            "slideSize": "ワイド スクリーン (16:9)",
            "designTheme": theme_meta["designTheme"],
            "variant": theme_meta["variant"],
            "colorTheme": theme_meta["colorTheme"],
            "fontTheme": theme_meta["fontTheme"],
        },
        "contentBlocks": {
            "slideMap": _p6_slide_map(spec),
            "preTaskState": {
                "documentInspected": False,
                "readOnlySet": False,
                "autoPresentationSet": False,
                "customShowCreated": False,
                "printSettingsApplied": False,
            },
        },
        "taskParams": _p6_task_params(spec),
    }


def build_project_json(project_id: int) -> dict:
    forbidden_map = {
        1: FORBIDDEN_P1, 2: FORBIDDEN_P2, 3: FORBIDDEN_P3,
        4: FORBIDDEN_P4, 5: FORBIDDEN_P5, 6: FORBIDDEN_P6,
        7: FORBIDDEN_P7, 8: FORBIDDEN_P8, 9: FORBIDDEN_P9, 10: FORBIDDEN_P10,
    }
    task_count_map = {1: 8, 2: 8, 3: 7, 4: 8, 5: 7, 6: 7, 7: 5, 8: 5, 9: 5, 10: 8}
    set_builders = {
        1: (build_project1_set, P1_SETS),
        2: (build_project2_set, P2_SETS),
        3: (build_project3_set, P3_SETS),
        4: (build_project4_set, P4_SETS),
        5: (build_project5_set, P5_SETS),
        6: (build_project6_set, P6_SETS),
        7: (build_project7_set, P7_SETS),
        8: (build_project8_set, P8_SETS),
        9: (build_project9_set, P9_SETS),
        10: (build_project10_set, P10_SETS),
    }
    forbidden = forbidden_map[project_id]
    task_count = task_count_map[project_id]
    builder, specs = set_builders[project_id]
    sets = []
    for i, theme_meta in enumerate(THEMES):
        set_no = i + 1
        sets.append(builder(set_no, theme_meta, specs[i]))
    textbook_defaults = {
        1: {"spacingPt": 4, "layoutInsert": "テーブルスライド", "layoutChange": "テキスト1スライド"},
        2: {"durationSec": 3, "autoAdvanceSec": 5, "transitionAll": "スプリット"},
        3: {"modelWidth": 2.5, "modelHeight": 6.5, "smartArtType": "タイムライン"},
        4: {
            "inputShapeType": "図形", "inputShapeColorDesc": "青", "textBoxShapeType": "テキストボックス",
            "shapeText": "教育者必見", "targetPhrase": "いつでも体験可能です!", "fillColor": "青、アクセント1",
            "glowPt": 18, "quickStyle": "楕円 ぼかし", "artEffect": "テクスチャライザー", "borderPt": 0.75,
        },
        5: {
            "alignShapeType": "丸", "rectShapeType": "四角", "sourceShapeType": "星",
            "targetShapeType": "スマイル", "gateShapeType": "論理積ゲート",
            "iconFillColor": "濃い赤", "slideTitleDecorative": "MOSって何？",
        },
        6: {
            "customShowName": "書式のポイント", "outlineCopies": 6,
            "notesCopies": 3, "handoutCopies": 4,
        },
        7: {
            "hyperlinkUrl": "https://rabbitway.jp/", "footerDomain": "rabbitway.jp",
            "commentText": "情報発信の責任を考える", "hyperlinkPhrase": "情報学習支援",
            "caseFooter": "参考事例", "outlineDoc": "まとめ.docx",
        },
        8: {
            "tableSlideTitle": "見せる！スライドの基本ルール",
            "tableStyle": "中間スタイル4-アクセント4",
            "backgroundColor": "濃い緑、テキスト2、白+基本色80％",
            "slideWidthCm": 27.54, "slideHeightCm": 17.46,
        },
        9: {
            "videoFile": "受講の様子.mp4", "videoSlideTitle": "スクールの様子",
            "trimStartSec": 4, "trimEndSec": 9, "fadeOutSec": 3,
            "categoryColumn": "結果", "dataColumn": "人数",
        },
        10: {
            "masterTheme": "木版活字", "customLayoutName": "タイトル付きの図と表",
            "handoutFooter": "四季を楽しむ",
        },
    }.get(project_id, {})
    per_set_must_vary = [
        "テーマ色とデザインテーマ",
        "スライドタイトルと本文",
        "数値パラメータ",
        "UI選択肢（レイアウト・画面切り替え・SmartArt等）",
    ]
    if project_id in (4, 5):
        per_set_must_vary.append("図形の形状（slideMap.shapes と問題文）")
    return {
        "variantRules": {
            "scope": (
                f"MOSスペシャリスト範疇。Project{project_id}の操作種別はPP修正版CSV準拠。"
                "数値・名称・文言・図形形状のみ差替。"
            ),
            "forbiddenTexts": forbidden,
            "perSetMustVary": per_set_must_vary,
            "mustNotReuseAcrossSets": [
                "spacingPt", "sectionNew", "zoomTitle", "transitionAll",
                "durationSec", "autoAdvanceSec", "smartArtColor",
                "inputShapeType", "alignShapeType", "customShowName",
            ],
            "uiPolicy": "ビルド2508のリボン表示名を使用",
            "specificity": "スライド番号・UI名・文字列・図形形状を完全一致で指定",
            "textbookDefaults": textbook_defaults,
        },
        "powerPointUiNames": POWERPOINT_UI,
        "sourceFile": f"Project{project_id}.pptx",
        "projectId": project_id,
        "variantType": "layout",
        "taskCount": task_count,
        "sets": sets,
    }


def write_all() -> None:
    out_dir = BASE / "類題Json"
    out_dir.mkdir(parents=True, exist_ok=True)
    for pid in range(1, 11):
        data = build_project_json(pid)
        path = out_dir / f"MOS_PowerPoint類題_project{pid}_配置別_5セット_問題文.json"
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"wrote {path.name}")


if __name__ == "__main__":
    write_all()
