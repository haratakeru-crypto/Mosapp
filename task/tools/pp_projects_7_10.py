"""PowerPoint variant builders for Projects 7-10."""

from __future__ import annotations

FORBIDDEN_P7 = [
    "情報学習支援", "rabbitway.jp", "情報発信の責任を考える",
    "参考事例", "事例１", "事例2", "まとめ.docx",
]
FORBIDDEN_P8 = [
    "見せる！", "見せる！スライドの基本ルール", "中間スタイル4-アクセント4",
    "濃い緑、テキスト2、白+基本色80％",
]
FORBIDDEN_P9 = [
    "英語教育", "スクールの様子", "受講の様子.mp4",
]
FORBIDDEN_P10 = [
    "木版活字", "四季を楽しむ", "タイトル付きの図と表", "メトロポリタン",
]

P7_INTROS = {
    1: "あなたは店舗運営における情報モラル研修資料を作成しています。",
    2: "あなたは医療機関の情報公開ガイドライン資料を作成しています。",
    3: "あなたは製造現場の情報セキュリティ啓発資料を作成しています。",
    4: "あなたは旅行会社の顧客情報保護研修資料を作成しています。",
    5: "あなたは学習塾の保護者向け情報モラル説明資料を作成しています。",
}

P7_SETS = [
    {
        "commentText": "投稿前の事実確認を徹底する",
        "hyperlinkPhrase": "店舗運営支援",
        "hyperlinkUrl": "https://cafe-guide.example.jp/",
        "outlineDoc": "まとめ_カフェ.docx",
        "footerDomain": "cafe-guide.example.jp",
        "caseFooter": "運営事例",
        "case1Title": "事例A",
        "case2Title": "事例B",
        "slideTitles": [
            "情報モラルの基本", "発信のルール", "個人情報の取扱", "事例の背景",
            "事例A", "事例B", "確認のポイント", "今後の取組",
        ],
    },
    {
        "commentText": "患者情報の取扱注意を周知する",
        "hyperlinkPhrase": "健診情報支援",
        "hyperlinkUrl": "https://clinic-info.example.jp/",
        "outlineDoc": "まとめ_医療.docx",
        "footerDomain": "clinic-info.example.jp",
        "caseFooter": "院内事例",
        "case1Title": "事例C",
        "case2Title": "事例D",
        "slideTitles": [
            "情報公開の基本", "記録の管理", "同意の取得", "事例の背景",
            "事例C", "事例D", "監査のポイント", "改善計画",
        ],
    },
    {
        "commentText": "設備データの持ち出し禁止を再確認",
        "hyperlinkPhrase": "保全情報支援",
        "hyperlinkUrl": "https://factory-safe.example.jp/",
        "outlineDoc": "まとめ_製造.docx",
        "footerDomain": "factory-safe.example.jp",
        "caseFooter": "安全事例",
        "case1Title": "事例E",
        "case2Title": "事例F",
        "slideTitles": [
            "情報セキュリティの基本", "アクセス制御", "ログの保管", "事例の背景",
            "事例E", "事例F", "点検のポイント", "教育計画",
        ],
    },
    {
        "commentText": "予約情報の第三者提供を禁止",
        "hyperlinkPhrase": "予約情報支援",
        "hyperlinkUrl": "https://tour-book.example.jp/",
        "outlineDoc": "まとめ_旅行.docx",
        "footerDomain": "tour-book.example.jp",
        "caseFooter": "対応事例",
        "case1Title": "事例G",
        "case2Title": "事例H",
        "slideTitles": [
            "顧客情報の基本", "予約データ管理", "クレーム対応", "事例の背景",
            "事例G", "事例H", "確認のポイント", "運用改善",
        ],
    },
    {
        "commentText": "成績情報の共有範囲を限定する",
        "hyperlinkPhrase": "学習記録支援",
        "hyperlinkUrl": "https://juku-portal.example.jp/",
        "outlineDoc": "まとめ_学習塾.docx",
        "footerDomain": "juku-portal.example.jp",
        "caseFooter": "指導事例",
        "case1Title": "事例I",
        "case2Title": "事例J",
        "slideTitles": [
            "保護者連絡の基本", "成績データ管理", "写真の掲載", "事例の背景",
            "事例I", "事例J", "確認のポイント", "年度計画",
        ],
    },
]

P8_INTROS = {
    1: "あなたはカフェの店舗プレゼン技法資料を作成しています。",
    2: "あなたは医療説明会のプレゼン技法資料を作成しています。",
    3: "あなたは工場見学会のプレゼン技法資料を作成しています。",
    4: "あなたは旅行商品説明のプレゼン技法資料を作成しています。",
    5: "あなたは学習塾説明会のプレゼン技法資料を作成しています。",
}

P8_SETS = [
    {
        "tableSlideTitle": "伝わるスライドの基本ルール",
        "tableStyle": "中間スタイル3-アクセント2",
        "backgroundColor": "濃い茶、テキスト2、白+基本色60％",
        "slideWidthCm": 25.4, "slideHeightCm": 19.05,
        "slideTitles": [
            "伝わるプレゼンの極意", "伝わるスライドの基本ルール",
            "視覚的な工夫", "構成のコツ", "まとめ",
        ],
    },
    {
        "tableSlideTitle": "説明会スライドの基本ルール",
        "tableStyle": "中間スタイル2-アクセント3",
        "backgroundColor": "濃い緑、テキスト2、白+基本色70％",
        "slideWidthCm": 26.0, "slideHeightCm": 18.0,
        "slideTitles": [
            "院内説明の極意", "説明会スライドの基本ルール",
            "図表の見せ方", "質疑の進め方", "まとめ",
        ],
    },
    {
        "tableSlideTitle": "見学会スライドの基本ルール",
        "tableStyle": "中間スタイル5-アクセント1",
        "backgroundColor": "濃い青灰、テキスト2、白+基本色75％",
        "slideWidthCm": 27.0, "slideHeightCm": 17.8,
        "slideTitles": [
            "見学会プレゼンの極意", "見学会スライドの基本ルール",
            "安全標識の見せ方", "質問対応のコツ", "まとめ",
        ],
    },
    {
        "tableSlideTitle": "提案スライドの基本ルール",
        "tableStyle": "中間スタイル4-アクセント5",
        "backgroundColor": "濃い水色、テキスト2、白+基本色65％",
        "slideWidthCm": 28.0, "slideHeightCm": 17.5,
        "slideTitles": [
            "提案プレゼンの極意", "提案スライドの基本ルール",
            "行程表の見せ方", "価格提示のコツ", "まとめ",
        ],
    },
    {
        "tableSlideTitle": "講座案内スライドの基本ルール",
        "tableStyle": "中間スタイル6-アクセント4",
        "backgroundColor": "濃いオレンジ、テキスト2、白+基本色55％",
        "slideWidthCm": 28.5, "slideHeightCm": 16.8,
        "slideTitles": [
            "講座説明の極意", "講座案内スライドの基本ルール",
            "進度表の見せ方", "保護者説明のコツ", "まとめ",
        ],
    },
]

P9_INTROS = {
    1: "あなたはカフェスタッフ研修プログラムの提案資料を作成しています。",
    2: "あなたは医療事務研修プログラムの提案資料を作成しています。",
    3: "あなたは設備保全研修プログラムの提案資料を作成しています。",
    4: "あなたは旅行代理店研修プログラムの提案資料を作成しています。",
    5: "あなたは講師研修プログラムの提案資料を作成しています。",
}

P9_SETS = [
    {
        "videoSlideTitle": "研修の様子",
        "videoFile": "接客研修.mp4",
        "trimStartSec": 3, "trimEndSec": 8,
        "fadeOutSec": 2,
        "categoryColumn": "項目", "dataColumn": "人数",
        "chartRows": [["抽出", 12], ["ラテ", 18], ["接客", 15]],
        "slideTitles": [
            "プログラム概要", "受講結果", "カリキュラム", "講師紹介", "研修の様子", "お問い合わせ",
        ],
    },
    {
        "videoSlideTitle": "実習の様子",
        "videoFile": "受付実習.mp4",
        "trimStartSec": 5, "trimEndSec": 11,
        "fadeOutSec": 3,
        "categoryColumn": "演習", "dataColumn": "人数",
        "chartRows": [["受付", 10], ["記録", 14], ["連絡", 11]],
        "slideTitles": [
            "研修概要", "受講結果", "演習項目", "講師紹介", "実習の様子", "お問い合わせ",
        ],
    },
    {
        "videoSlideTitle": "保全実習の様子",
        "videoFile": "点検実習.mp4",
        "trimStartSec": 5, "trimEndSec": 10,
        "fadeOutSec": 4,
        "categoryColumn": "工程", "dataColumn": "人数",
        "chartRows": [["点検", 9], ["修理", 13], ["記録", 12]],
        "slideTitles": [
            "保全研修概要", "受講結果", "実習内容", "講師紹介", "保全実習の様子", "お問い合わせ",
        ],
    },
    {
        "videoSlideTitle": "接客研修の様子",
        "videoFile": "予約対応.mp4",
        "trimStartSec": 2, "trimEndSec": 7,
        "fadeOutSec": 3,
        "categoryColumn": "区分", "dataColumn": "人数",
        "chartRows": [["予約", 11], ["変更", 9], ["案内", 16]],
        "slideTitles": [
            "代理店研修概要", "受講結果", "演習メニュー", "講師紹介", "接客研修の様子", "お問い合わせ",
        ],
    },
    {
        "videoSlideTitle": "授業実習の様子",
        "videoFile": "模擬授業.mp4",
        "trimStartSec": 6, "trimEndSec": 11,
        "fadeOutSec": 5,
        "categoryColumn": "科目", "dataColumn": "人数",
        "chartRows": [["数学", 8], ["英語", 10], ["理科", 7]],
        "slideTitles": [
            "講師研修概要", "受講結果", "演習計画", "講師紹介", "授業実習の様子", "お問い合わせ",
        ],
    },
]

P10_INTROS = {
    1: "あなたはカフェの季節メニュー紹介資料を作成しています。",
    2: "あなたは院内四季の健康案内資料を作成しています。",
    3: "あなたは工場の季節保全計画資料を作成しています。",
    4: "あなたは旅行の四季プラン紹介資料を作成しています。",
    5: "あなたは学習塾の四季講座案内資料を作成しています。",
}

P10_SETS = [
    {
        "masterTheme": "イオン",
        "customLayoutName": "図表カスタム_カフェ",
        "handoutFooter": "季節メニューを楽しむ",
        "slideTitles": ["春のメニュー", "夏のドリンク", "秋の限定", "冬の温かい一杯", "豆の産地", "焙煎の違い", "まとめ"],
    },
    {
        "masterTheme": "センチュリー",
        "customLayoutName": "図表カスタム_医療",
        "handoutFooter": "健康な四季を送る",
        "slideTitles": ["春の健診", "夏の熱中症", "秋の免疫", "冬の感染症", "食事の工夫", "運動の習慣", "まとめ"],
    },
    {
        "masterTheme": "バーチ",
        "customLayoutName": "図表カスタム_製造",
        "handoutFooter": "安全な運転を続ける",
        "slideTitles": ["春の点検", "夏の冷却", "秋の整備", "冬の暖機", "潤滑管理", "異常記録", "まとめ"],
    },
    {
        "masterTheme": "ロード",
        "customLayoutName": "図表カスタム_旅行",
        "handoutFooter": "旅の四季を味わう",
        "slideTitles": ["春の桜旅", "夏の海旅", "秋の紅葉", "冬の温泉", "交通手段", "宿泊の選び方", "まとめ"],
    },
    {
        "masterTheme": "ダイナミック",
        "customLayoutName": "図表カスタム_学習塾",
        "handoutFooter": "学びの四季を深める",
        "slideTitles": ["春期講座", "夏期講座", "秋期講座", "冬期講座", "進度管理", "自習支援", "まとめ"],
    },
]


def _p7_slide_map(spec: dict) -> list[dict]:
    sm = []
    case1, case2 = spec["case1Title"], spec["case2Title"]
    for i, title in enumerate(spec["slideTitles"], 1):
        e: dict = {"slideNo": i, "layout": "タイトルとコンテンツ", "title": title}
        if i == 1:
            e["layout"] = "タイトルスライド"
            e["objects"] = [spec["hyperlinkPhrase"]]
            e["notes"] = f"7-1: コメント「{spec['commentText']}」未投稿 / 7-2: ハイパーリンク未設定"
        elif i == 5:
            e["notes"] = f"7-5: フッター「{spec['caseFooter']}」未設定（事例スライド）"
        elif i == 6:
            e["notes"] = f"7-5: フッター「{spec['caseFooter']}」未設定 / 7-3: この後にアウトライン挿入予定"
        elif i == 8:
            e["notes"] = f"7-3: スライド6の後に「{spec['outlineDoc']}」から挿入する想定（未挿入）"
        sm.append(e)
    return sm


def _p7_tasks(spec: dict) -> list[str]:
    return [
        f"タスク7-1　スライド１にコメント「\"{spec['commentText']}\"」を挿入します。",
        f"タスク7-2　スライド１枚目の文字列「{spec['hyperlinkPhrase']}」に、"
        f"Webページ「\"{spec['hyperlinkUrl']}\"」を表示するハイパーリンクを設定します。",
        f"タスク7-3　スライド６の後ろに、文書「{spec['outlineDoc']}」のアウトラインを使用してスライドを挿入します。",
        f"タスク7-4　スライドのフッターに、スライド番号と「\"{spec['footerDomain']}\"」をタイトルスライド以外に追加します。",
        f"タスク7-5　５枚目の「{spec['case1Title']}」と６枚目の「{spec['case2Title']}」のフッターに"
        f"「\"{spec['caseFooter']}\"」と挿入します。他のスライドには表示しません。",
    ]


def _p7_task_params(spec: dict) -> dict:
    return {
        "task7_1": {"slideNo": 1, "commentText": spec["commentText"], "operation": "insertComment"},
        "task7_2": {
            "slideNo": 1,
            "hyperlinkPhrase": spec["hyperlinkPhrase"],
            "hyperlinkUrl": spec["hyperlinkUrl"],
            "operation": "insertHyperlink",
        },
        "task7_3": {
            "insertAfterSlide": 6,
            "outlineDoc": spec["outlineDoc"],
            "operation": "insertOutlineSlides",
        },
        "task7_4": {
            "footerDomain": spec["footerDomain"],
            "excludeTitleSlide": True,
            "slideNumber": True,
            "operation": "headerFooterAll",
        },
        "task7_5": {
            "slideNos": [5, 6],
            "slideTitles": [spec["case1Title"], spec["case2Title"]],
            "footerText": spec["caseFooter"],
            "operation": "headerFooterSelected",
        },
    }


def _p8_slide_map(spec: dict) -> list[dict]:
    tbl_title = spec["tableSlideTitle"]
    sm = []
    for i, title in enumerate(spec["slideTitles"], 1):
        e: dict = {"slideNo": i, "layout": "タイトルとコンテンツ", "title": title}
        if i == 1:
            e["layout"] = "タイトルスライド"
            e["notes"] = f"8-2: 背景色「{spec['backgroundColor']}」未設定"
        elif title == tbl_title:
            e["objects"] = ["表", "イラスト"]
            e["notes"] = f"8-1: 表スタイル「{spec['tableStyle']}」未変更 / 8-5: イラストのグレースケール未設定"
        sm.append(e)
    return sm


def _p8_tasks(spec: dict) -> list[str]:
    tbl = spec["tableSlideTitle"]
    return [
        f"タスク8-1　スライド「{tbl}」の表を、スタイル「{spec['tableStyle']}」に変更します。"
        "表スタイルのオプションを設定して、行が１行おきに変更されないようにします。",
        f"タスク8-2　スライド１の背景を「{spec['backgroundColor']}」に変更します。",
        "タスク8-3　スライドのサイズを「画面に合わせる16：9」にします。",
        f"タスク8-4　スライドの大きさを、高さ「\"{spec['slideHeightCm']}\"cm」、"
        f"幅「\"{spec['slideWidthCm']}\"cm」に変更します。スライドは画面に合わせます。",
        "タスク8-5　スライドをグレースケールで表示し、スライド２のイラストを「明るいグレースケール」にします。",
    ]


def _p8_task_params(spec: dict) -> dict:
    return {
        "task8_1": {
            "slideTitle": spec["tableSlideTitle"],
            "tableStyle": spec["tableStyle"],
            "bandedRows": False,
            "operation": "tableStyle",
        },
        "task8_2": {
            "slideNo": 1,
            "backgroundColor": spec["backgroundColor"],
            "operation": "slideBackground",
        },
        "task8_3": {"slideSize": "画面に合わせる(16:9)", "operation": "slideSizePreset"},
        "task8_4": {
            "widthCm": spec["slideWidthCm"],
            "heightCm": spec["slideHeightCm"],
            "scaleToFit": True,
            "operation": "slideSizeCustom",
        },
        "task8_5": {
            "slideNo": 2,
            "grayscaleView": True,
            "illustrationStyle": "明るいグレースケール",
            "operation": "grayscaleIllustration",
        },
    }


def _p9_slide_map(spec: dict) -> list[dict]:
    sm = []
    for i, title in enumerate(spec["slideTitles"], 1):
        e: dict = {"slideNo": i, "layout": "タイトルとコンテンツ", "title": title}
        if i == 1:
            e["layout"] = "タイトルスライド"
            e["objects"] = ["オーディオ"]
            e["notes"] = f"9-3: オーディオのスライド切替後再生・フェード{spec['fadeOutSec']}秒未設定"
        elif i == 2:
            e["objects"] = ["表"]
            e["notes"] = f"9-4/9-5: グラフ「{spec['categoryColumn']}」×「{spec['dataColumn']}」未作成・データテーブル未設定"
        elif title == spec["videoSlideTitle"]:
            e["objects"] = [spec["videoFile"]]
            e["notes"] = (
                f"9-1: ビデオ「{spec['videoFile']}」未挿入 / "
                f"9-2: トリミング {spec['trimStartSec']}〜{spec['trimEndSec']}秒未設定"
            )
        sm.append(e)
    return sm


def _p9_tasks(spec: dict) -> list[str]:
    vt = spec["videoSlideTitle"]
    return [
        f"タスク9-1　スライド「{vt}」に、動画「{spec['videoFile']}」を挿入します。挿入にはアイコンを使います。",
        f"タスク9-2　スライド５のビデオを、開始を「\"{spec['trimStartSec']}\"秒」、"
        f"終了を「\"{spec['trimEndSec']}\"秒」に設定します。",
        f"タスク9-3　スライド１のオーディオを、スライドを切り替えても1回だけ再生するように設定します。"
        f"再生は\"{spec['fadeOutSec']}\"秒かけてフェードアウトするようにします。",
        f"タスク9-4　スライド２のプレースホルダーに表の内容を表す「集合縦棒」グラフを作成します。"
        f"「{spec['categoryColumn']}」の列を項目、「{spec['dataColumn']}」の列をデータ系列として使用します。"
        "表のデータはグラフシートにコピーまたは手入力でもかまいません。",
        "タスク9-5　スライド２のグラフに［凡例マーカーなし］のデータテーブルを表示し、タイトルと凡例を削除します。",
    ]


def _p9_task_params(spec: dict) -> dict:
    return {
        "task9_1": {
            "slideTitle": spec["videoSlideTitle"],
            "videoFile": spec["videoFile"],
            "operation": "insertVideo",
        },
        "task9_2": {
            "slideNo": 5,
            "trimStartSec": spec["trimStartSec"],
            "trimEndSec": spec["trimEndSec"],
            "operation": "trimVideo",
        },
        "task9_3": {
            "slideNo": 1,
            "playAcrossSlides": True,
            "fadeOutSec": spec["fadeOutSec"],
            "operation": "audioSettings",
        },
        "task9_4": {
            "slideNo": 2,
            "chartType": "集合縦棒",
            "categoryColumn": spec["categoryColumn"],
            "dataColumn": spec["dataColumn"],
            "tableRows": spec["chartRows"],
            "operation": "insertChart",
        },
        "task9_5": {
            "slideNo": 2,
            "dataTable": "凡例マーカーなし",
            "removeTitle": True,
            "removeLegend": True,
            "operation": "chartDataTable",
        },
    }


def _p10_slide_map(spec: dict) -> list[dict]:
    sm = []
    for i, title in enumerate(spec["slideTitles"], 1):
        e: dict = {"slideNo": i, "layout": "タイトルとコンテンツ", "title": title}
        if i == 1:
            e["layout"] = "タイトルスライド"
        if i == 2:
            e["layout"] = "2 つのコンテンツ"
            e["notes"] = "10-4: 2つのコンテンツの背景デザイン非表示未実施"
        sm.append(e)
    return sm


def _p10_tasks(spec: dict) -> list[str]:
    return [
        f"タスク10-1　スライドマスターにテーマ「{spec['masterTheme']}」を設定します。",
        "タスク10-2　スライドマスターにスライド番号を挿入します。",
        "タスク10-3　スライドマスターの［タイトルスライド］レイアウトのスライド番号を非表示にします。",
        "タスク10-4　スライドマスターの［２つのコンテンツ］レイアウトのスライドの背景のデザインを非表示にします。",
        "タスク10-5　スライドマスターで［フッター］のプレースホルダーを削除します。",
        f"タスク10-6　スライドマスターの［タイトルのみ］レイアウトをもとに「\"{spec['customLayoutName']}\"」の名前で"
        "レイアウトを作成します。図のプレースホルダーをスライドの左側、表のプレースホルダーを右側に配置します。",
        "タスク10-7　配布資料マスターの日付を削除します。",
        f"タスク10-8　配布資料マスターのフッターに「\"{spec['handoutFooter']}\"」と表示します。",
    ]


def _p10_task_params(spec: dict) -> dict:
    return {
        "task10_1": {"masterTheme": spec["masterTheme"], "operation": "masterTheme"},
        "task10_2": {"operation": "masterSlideNumber"},
        "task10_3": {"layout": "タイトルスライド", "operation": "hideMasterSlideNumber"},
        "task10_4": {"layout": "2 つのコンテンツ", "operation": "hideMasterBackground"},
        "task10_5": {"operation": "deleteMasterFooter"},
        "task10_6": {
            "baseLayout": "タイトルのみ",
            "customLayoutName": spec["customLayoutName"],
            "operation": "customMasterLayout",
        },
        "task10_7": {"operation": "deleteHandoutDate"},
        "task10_8": {"handoutFooter": spec["handoutFooter"], "operation": "handoutFooter"},
    }


def build_project7_set(set_no: int, theme_meta: dict, spec: dict) -> dict:
    return {
        "setNo": set_no,
        "theme": theme_meta["theme"],
        "workbook": f"MOS_PowerPoint類題_project7_配置別_セット{set_no}_{theme_meta['theme']}.pptx",
        "problemStatement": P7_INTROS[set_no],
        "tasks": _p7_tasks(spec),
        "layout": {
            "slideCount": len(spec["slideTitles"]),
            "slideSize": "ワイド スクリーン (16:9)",
            "designTheme": theme_meta["designTheme"],
            "variant": theme_meta["variant"],
            "colorTheme": theme_meta["colorTheme"],
            "fontTheme": theme_meta["fontTheme"],
        },
        "contentBlocks": {
            "slideMap": _p7_slide_map(spec),
            "preTaskState": {
                "commentOnSlide1": False,
                "hyperlinkOnSlide1": False,
                "outlineInserted": False,
                "globalFooterApplied": False,
                "caseFooterApplied": False,
            },
        },
        "taskParams": _p7_task_params(spec),
    }


def build_project8_set(set_no: int, theme_meta: dict, spec: dict) -> dict:
    return {
        "setNo": set_no,
        "theme": theme_meta["theme"],
        "workbook": f"MOS_PowerPoint類題_project8_配置別_セット{set_no}_{theme_meta['theme']}.pptx",
        "problemStatement": P8_INTROS[set_no],
        "tasks": _p8_tasks(spec),
        "layout": {
            "slideCount": len(spec["slideTitles"]),
            "slideSize": "ワイド スクリーン (16:9)",
            "designTheme": theme_meta["designTheme"],
            "variant": theme_meta["variant"],
            "colorTheme": theme_meta["colorTheme"],
            "fontTheme": theme_meta["fontTheme"],
        },
        "contentBlocks": {
            "slideMap": _p8_slide_map(spec),
            "preTaskState": {
                "tableStyleApplied": False,
                "backgroundOnSlide1": False,
                "slideSize169": False,
                "slideSizeCustom": False,
                "grayscaleIllustration": False,
            },
        },
        "taskParams": _p8_task_params(spec),
    }


def build_project9_set(set_no: int, theme_meta: dict, spec: dict) -> dict:
    return {
        "setNo": set_no,
        "theme": theme_meta["theme"],
        "workbook": f"MOS_PowerPoint類題_project9_配置別_セット{set_no}_{theme_meta['theme']}.pptx",
        "problemStatement": P9_INTROS[set_no],
        "tasks": _p9_tasks(spec),
        "layout": {
            "slideCount": len(spec["slideTitles"]),
            "slideSize": "ワイド スクリーン (16:9)",
            "designTheme": theme_meta["designTheme"],
            "variant": theme_meta["variant"],
            "colorTheme": theme_meta["colorTheme"],
            "fontTheme": theme_meta["fontTheme"],
        },
        "contentBlocks": {
            "slideMap": _p9_slide_map(spec),
            "chartTableData": {
                "headers": [spec["categoryColumn"], spec["dataColumn"]],
                "rows": spec["chartRows"],
            },
            "preTaskState": {
                "videoInserted": False,
                "videoTrimmed": False,
                "audioConfigured": False,
                "chartCreated": False,
                "chartDataTable": False,
            },
        },
        "taskParams": _p9_task_params(spec),
    }


def build_project10_set(set_no: int, theme_meta: dict, spec: dict) -> dict:
    return {
        "setNo": set_no,
        "theme": theme_meta["theme"],
        "workbook": f"MOS_PowerPoint類題_project10_配置別_セット{set_no}_{theme_meta['theme']}.pptx",
        "problemStatement": P10_INTROS[set_no],
        "tasks": _p10_tasks(spec),
        "layout": {
            "slideCount": len(spec["slideTitles"]),
            "slideSize": "ワイド スクリーン (16:9)",
            "designTheme": theme_meta["designTheme"],
            "variant": theme_meta["variant"],
            "colorTheme": theme_meta["colorTheme"],
            "fontTheme": theme_meta["fontTheme"],
        },
        "contentBlocks": {
            "slideMap": _p10_slide_map(spec),
            "preTaskState": {
                "masterThemeSet": False,
                "masterSlideNumber": False,
                "titleLayoutNumberHidden": False,
                "twoContentBgHidden": False,
                "masterFooterDeleted": False,
                "customLayoutCreated": False,
                "handoutDateDeleted": False,
                "handoutFooterSet": False,
            },
        },
        "taskParams": _p10_task_params(spec),
    }
