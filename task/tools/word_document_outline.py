"""Build documentOutline (report flesh) for Word variant sets P6-P10."""

from __future__ import annotations

from typing import Any


def _sec(level: int, text: str, body: list[str] | None = None, children: list | None = None, note: str = "") -> dict:
    d: dict[str, Any] = {"level": level, "text": text, "body": body or []}
    if children:
        d["children"] = children
    if note:
        d["note"] = note
    return d


def _depth4_chain(h2: str, h3: str, h4: str, bodies: list[str]) -> list[dict]:
    return [
        _sec(
            2,
            h2,
            bodies[:1],
            [
                _sec(
                    3,
                    h3,
                    bodies[1:2] if len(bodies) > 1 else [],
                    [_sec(4, h4, bodies[2:] if len(bodies) > 2 else bodies[-1:])],
                )
            ],
        )
    ]


def build_outline(project_id: int, s: dict) -> dict:
    theme = s["theme"]
    ly = s["layout"]
    cb = s.get("contentBlocks", {})
    headings = ly.get("headings", [])
    h1, h2, h3 = (headings + ["", "", ""])[:3]

    if project_id == 6:
        p1h = s["taskParams"]["task6_1"]["heading"]
        p4h = s["taskParams"]["task6_4"]["heading"]
        titles = {
            "カフェ": "駅前店 月次売上報告書（2026年4月）",
            "医療": "本院 月次外来報告書（2026年4月）",
            "製造": "第一工場 月次出荷報告書（2026年4月）",
            "旅行": "国内ツアー 月次実績報告書（2026年4月）",
            "学習塾": "本校 月次受講報告書（2026年4月）",
        }
        return {
            "title": titles.get(theme, f"{theme} 月次報告書"),
            "minHeadingLevel": 4,
            "sections": [
                _sec(
                    1,
                    h1 or "（1）概要",
                    cb.get("introParagraphs", []),
                    _depth4_chain(
                        "1.1 報告の目的",
                        "1.1.1 集計範囲",
                        "1.1.1.1 対象期間",
                        [
                            f"{theme}テーマの月次実績を関係者に共有する。",
                            "対象は当月1日から末日までのデータです。",
                            "4月分を集計し、前月・前年同月と比較します。",
                        ],
                    ),
                ),
                _sec(
                    1,
                    p1h,
                    [f"以下は{p1h}の原データです（表変換前）。"],
                    _depth4_chain(
                        "2.1 カテゴリ別所感",
                        "2.1.1 好調項目",
                        "2.1.1.1 背景",
                        ["主力カテゴリが前年を上回りました。", "週末の来店が増加傾向です。", "天候要因の影響は限定的でした。"],
                    ),
                    note="6-1: 直下にタブ区切りテキスト（未表化）",
                ),
                _sec(
                    1,
                    p4h,
                    [f"{p4h}の数値は下表のとおりです。"],
                    _depth4_chain(
                        "3.1 月次比較",
                        "3.1.1 増減要因",
                        "3.1.1.1 施策効果",
                        ["前月比で伸長した項目があります。", "販促施策の効果が表れています。", "来月も同水準を目指します。"],
                    ),
                    note="6-4: 直下に売上表（列幅不均等・1行目効果なし）",
                ),
                _sec(
                    1,
                    h3 or "（3）今後の施策",
                    ["来月に向けた改善点をまとめます。"],
                    _depth4_chain(
                        "4.1 重点施策",
                        "4.1.1 実施予定",
                        "4.1.1.1 スケジュール",
                        ["重点施策を2件実施予定です。", "担当者へ周知済みです。", "第2週から順次開始します。"],
                    ),
                ),
            ],
        }

    if project_id == 7:
        titles = {
            "カフェ": "焙煎体験会 ご案内",
            "医療": "健康講座 ご案内",
            "製造": "安全講習会 ご案内",
            "旅行": "添乗員研修 ご案内",
            "学習塾": "保護者説明会 ご案内",
        }
        intro = cb.get("introParagraphs", [])
        body = cb.get("bodyParagraphs", [])
        return {
            "title": titles.get(theme, f"{theme} ご案内"),
            "minHeadingLevel": 4,
            "sections": [
                _sec(1, h1, intro, _depth4_chain("1.1 開催趣旨", "1.1.1 対象者", "1.1.1.1 参加資格", body[:3] or ["ご参加をお待ちしています。"])),
                _sec(1, h2, body[3:4] if len(body) > 3 else ["詳細は以下のとおりです。"], _depth4_chain("2.1 当日の流れ", "2.1.1 受付", "2.1.1.1 開始時刻", body[:4] if body else ["受付は開始15分前です。"])),
                _sec(
                    1,
                    h3,
                    ["お申し込み方法をご確認ください。"],
                    _depth4_chain("3.1 申込方法", "3.1.1 連絡先", "3.1.1.1 受付時間", ["電話またはWebから申込可能です。", "定員に達し次第締切ます。", "平日9時から17時まで受付。"]),
                    note="ヘッダー・会社名プロパティは未設定（7-2/7-3は受験者操作）",
                ),
                _sec(1, "お問い合わせ", ["ご不明点は事務局までご連絡ください。"], _depth4_chain("4.1 担当窓口", "4.1.1 電話番号", "4.1.1.1 担当者", ["平日のみ対応いたします。", "メールでも受付可能です。", "担当：イベント係"])),
            ],
        }

    if project_id == 8:
        sub = cb.get("subtitle", "")
        bib_h = s["taskParams"]["task8_4"]["heading"]
        fn_h = s["taskParams"]["task8_2"]["heading"]
        body_map = cb.get("bodyUnderHeadings", {})
        fn_anchor = s["taskParams"]["task8_2"]["footnoteAnchor"]
        return {
            "title": f"{theme}セキュリティレポート",
            "subtitle": sub,
            "minHeadingLevel": 4,
            "sections": [
                _sec(
                    1,
                    "はじめに",
                    [f"本レポートは{theme}分野のセキュリティ対策を整理します。"],
                    _depth4_chain("1.1 背景", "1.1.1 リスク概要", "1.1.1.1 想定脅威", ["情報漏えいへの備えが重要です。", "内部不正と外部攻撃の両面に対応します。", "運用ルールの徹底が前提です。"]),
                ),
                _sec(
                    1,
                    fn_h,
                    body_map.get(fn_h, []),
                    _depth4_chain(
                        "2.1 基本概念",
                        "2.1.1 用語整理",
                        f"2.1.1.1 {fn_anchor}の位置づけ",
                        [f"「{fn_anchor}」は本文中で重要語として扱います。", "脚注は未挿入です。", "関連規格に沿って説明します。"],
                    ),
                    note="8-1: 目次未挿入（図形内・副題下以外の指定位置） / 8-2: 脚注未挿入",
                ),
                _sec(
                    1,
                    headings[1] if len(headings) > 1 else "対策の基本",
                    body_map.get(headings[1] if len(headings) > 1 else "", []),
                    _depth4_chain("3.1 運用上の注意", "3.1.1 日常点検", "3.1.1.1 記録保管", ["ログを定期確認します。", "異常時は即時報告します。", "記録は90日間保管します。"]),
                ),
                _sec(
                    1,
                    bib_h,
                    ["以下に参考資料を列挙します。"],
                    _depth4_chain("4.1 参照ガイド", "4.1.1 社内規程", "4.1.1.1 改訂履歴", ["最新版を参照してください。", "年1回見直しを実施します。", "特殊文字は未挿入です。"]),
                    note=f"8-4: 見出し「{bib_h}」先頭に特殊文字未挿入 / 8-5: SmartArt未挿入 / 8-7: リンク未設定",
                ),
            ],
        }

    if project_id == 9:
        ns = cb.get("numberedSection", {})
        ns_h = ns.get("heading", headings[0] if headings else "")
        t2h = s["taskParams"]["task9_2"]["heading"]
        t3h = s["taskParams"]["task9_3"]["tableHeading"]
        p2 = s["taskParams"]["task9_2"]
        p3 = s["taskParams"]["task9_3"]
        p4 = s["taskParams"]["task9_4"]
        sep_note = p2.get("separatorLabel", "コンマ区切り")
        margin_note = f"セル{p3.get('marginSideLabel', '右')}余白は未調整です"
        sort_note = f"「{p4['sortColumn']}」の{p4.get('sortOrderLabel', '降順')}で並べ替えます"
        return {
            "title": f"{theme} 催事売上報告書",
            "minHeadingLevel": 4,
            "sections": [
                _sec(
                    1,
                    ns_h,
                    [f"{ns_h}の計画内容です。"],
                    _depth4_chain("1.1 企画概要", "1.1.1 実施時期", "1.1.1.1 担当部署", ["季節イベントを企画しました。", "販促と連動して実施します。", "営業部が主担当です。"]),
                    note="9-1: 番号付き段落「" + ns.get("paragraphs", [""])[-2] + "」等を配置",
                ),
                _sec(1, t2h, ["月別の実績データです。"], note=f"9-2: 直下に表（{sep_note}変換前）"),
                _sec(
                    1,
                    t3h,
                    ["店舗別・拠点別の売上です。"],
                    _depth4_chain("3.1 分析メモ", "3.1.1 ソート前提", "3.1.1.1 集計単位", [sort_note + "。", margin_note + "。", "単位は千円です。"]),
                    note="9-3/9-4: 表はセル余白・並べ替え未実施 / 9-5: 変更履歴は未処理のまま",
                ),
                _sec(
                    1,
                    headings[-1] if headings else "まとめ",
                    ["今後の改善点を記載します。"],
                    _depth4_chain("4.1 次回施策", "4.1.1 予算", "4.1.1.1 承認状況", ["次回は規模を拡大予定です。", "予算案は上長確認中です。", "承認後に周知します。"]),
                ),
            ],
        }

    if project_id == 10:
        h2 = s["taskParams"]["task10_2"]["heading"]
        h3 = s["taskParams"]["task10_3"]["heading"]
        p3 = s["taskParams"]["task10_3"]
        p4 = s["taskParams"]["task10_4"]
        p5 = s["taskParams"]["task10_5"]
        search = s["taskParams"]["task10_1"]["searchText"]
        bodies = cb.get("bodyWithSearchTerms", [])
        s2 = cb.get("section10_2_paragraphs", [])
        s3 = cb.get("section10_3_paragraphs", [])
        last = cb.get("lastTwoLines", [])
        lst = p5["lineSpacingType"]
        val = p5.get("lineSpacingValue")
        if lst in ("1行", "1.5行", "2行"):
            ls_note = lst
        elif lst == "倍数":
            ls_note = f"倍数{val}"
        else:
            ls_note = f"{lst}{val}pt"
        if p4["formatType"] == "bold":
            fmt_note = f"「{p4['formatTarget']}」は太字未設定"
        elif p4["formatType"] == "color":
            fmt_note = f"「{p4['formatTarget']}」の色は{p4['formatColor']}未設定"
        else:
            fmt_note = f"「{p4['formatTarget']}」のサイズは{p4['formatSizePt']}pt未設定"
        return {
            "title": f"{theme}向け IT安全啓発レポート",
            "minHeadingLevel": 4,
            "sections": [
                _sec(
                    1,
                    headings[0] if headings else "はじめに",
                    bodies[:2] if bodies else [f"「{search}」を含む本文を配置します。"],
                    _depth4_chain("1.1 背景", "1.1.1 リスク", "1.1.1.1 影響範囲", bodies[2:5] if len(bodies) > 2 else ["日々の業務で注意が必要です。"]),
                    note=f"10-1: 「{search}」は置換前のまま複数回出現",
                ),
                _sec(1, h2, s2[:1] if s2 else [], _depth4_chain("2.1 基本方針", "2.1.1 実践項目", "2.1.1.1 頻度", s2 or ["端末管理を徹底します。"]), note="10-2: 行頭文字（画像）は未設定"),
                _sec(1, h3, s3[:1] if s3 else [], _depth4_chain("3.1 補足事項", "3.1.1 留意点", "3.1.1.1 連絡先", s3 or ["追加の注意事項です。"]), note=f"10-3: {p3['bulletFont']} {p3['bulletCharCode']}は未設定"),
                _sec(
                    1,
                    "まとめ",
                    last,
                    _depth4_chain("4.1 振り返り", "4.1.1 次回予定", "4.1.1.1 資料更新", ["継続的な啓発を行います。", "次回は運用例を追加します。", f"用紙{p5['paperSize']}・行間{ls_note}は未設定（10-5）"]),
                    note=f"10-4: {fmt_note} / 10-5: 末尾2行行間{ls_note}未設定",
                ),
            ],
        }

    return {"title": theme, "minHeadingLevel": 4, "sections": []}


def render_outline_md(outline: dict, indent: int = 0) -> list[str]:
    lines: list[str] = []
    prefix = "  " * indent
    if indent == 0 and outline.get("title"):
        lines.append(f"**文書タイトル:** {outline['title']}")
        if outline.get("subtitle"):
            lines.append(f"**副題:** {outline['subtitle']}")
        lines.append(f"**見出し階層:** 見出し1〜{outline.get('minHeadingLevel', 4)}まで（最低1系統）")
        lines.append("")
    for sec in outline.get("sections", []):
        lv = sec["level"]
        lines.append(f"{prefix}- 見出し{lv}: {sec['text']}")
        if sec.get("note"):
            lines.append(f"{prefix}  - ※ {sec['note']}")
        for p in sec.get("body", []):
            lines.append(f"{prefix}  - 本文: {p}")
        for child in sec.get("children", []):
            lines.extend(render_outline_md({"sections": [child]}, indent + 1))
        lines.append("")
    return lines


def apply_outlines_to_json(data: dict) -> dict:
    pid = data["projectId"]
    if pid not in (6, 7, 8, 9, 10):
        return data
    for s in data["sets"]:
        outline = build_outline(pid, s)
        s.setdefault("contentBlocks", {})["documentOutline"] = outline
    return data


if __name__ == "__main__":
    import json
    from pathlib import Path

    base = Path(__file__).resolve().parents[1]
    for pid in range(6, 11):
        path = base / "類題Json" / f"MOS_Word類題_project{pid}_配置別_5セット_問題文.json"
        data = json.loads(path.read_text(encoding="utf-8"))
        data = apply_outlines_to_json(data)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"updated outlines: project {pid}")
