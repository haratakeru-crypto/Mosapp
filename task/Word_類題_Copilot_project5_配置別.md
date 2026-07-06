# Word 類題作成 — Copilot 用プロンプト（Project5 配置別）

Microsoft Copilot に **コピペして使う** プロンプト集です。
project5 の類題文書（5セット×8タスク）を作成します。

**添付ファイル（推奨）:**
- `@task/Word問題文一覧_操作手順.csv`（projectId=5 の行）

**問題文ファイル（完成初稿）:**
- [`task/MOS_Word類題_project5_配置別_5セット_問題文.md`](MOS_Word類題_project5_配置別_5セット_問題文.md)
- [`task/類題Json/MOS_Word類題_project5_配置別_5セット_問題文.json`](類題Json/MOS_Word類題_project5_配置別_5セット_問題文.json)
- [`task/類題Json/MOS_Word類題_project5_配置別_5セット_問題文_アプリ用.json`](類題Json/MOS_Word類題_project5_配置別_5セット_問題文_アプリ用.json)

**禁止文字列（教材流用不可）:** `5月21日より5日間` `TOEICテスト対策セミナー` `TOEIC` `セミナー案内`

**タスク5-1〜5-8 のパラメータ（セットごとに異なる）:**

| セット | 5-1挿入/折返 | 5-2折返 | 5-5位置 | 5-6表現 | 5-7表現 | 5-8操作 |
|--------|-------------|--------|--------|--------|--------|--------|
| 1 カフェ | paragraphStart / 四角形 | 貫通 | titleLeft | 代替テキスト | 装飾化 | removeBackgroundOnly |
| 2 医療 | paragraphEnd / 貫通 | 四角形 | underTitle | 代替テキスト | 装飾化 | removeBackgroundOnly |
| 3 製造 | underHeading / 狭く | 背面 | titleRight | 説明 | screenReaderHidden | addForegroundMark |
| 4 旅行 | besideParagraph / 前面 | 上下 | headingRight | 説明 | screenReaderHidden | addForegroundMark |
| 5 学習塾 | underTitle / 背面 | 前面 | besideTitle | 説明 | screenReaderHidden | addForegroundMark |

**注意:** 5-1折返しは上下以外。5-2は狭く以外。5-3は水彩：スポンジ以外。
5-4は25pt以外。5-5はハードエッジ以外。5-6は3セット以上で「代替テキスト」不使用。
5-7は3セット以上でスクリーンリーダー非表示表現。5-8は3セット以上で前景領域追加。

**Word UI パス（ビルド2508）:**

- 画像挿入: 挿入 → 画像
- 折り返し: 図の形式 → 文字列の折り返し
- 代替テキスト: 図の形式 → 代替テキスト
- 背景削除: 図の形式 → 図の効果 → 背景の削除

---

## 0. 使い方（3セッション × 1セット）

```
1チャット = 1セット（§1〜§5 のいずれか1節のみ使用）
      ↓
セッション1 → 該当 § の「セッション1」プロンプト
      ↓  docx を保存
セッション2 → 同 § の「セッション2」プロンプト（docx 添付）
      ↓
セッション3 → 同 § の「セッション3」プロンプト（検証済みパラメータ表を貼る）
```

- **問題文はセッション3まで書かない**（セッション1はレイアウト仕様のみ）
- **操作の大分類を変更しない**（5-1=画像挿入、5-2=折り返し、5-3=アート効果、5-4=ぼかし、5-5=面取り、5-6=代替テキスト、5-7=装飾化、5-8=背景削除）
- 文書名: `MOS_Word類題_project5_配置別_セット{N}_{テーマ}.docx`

---

## §1 セット1 — カフェ

**文書名:** `MOS_Word類題_project5_配置別_セット1_カフェ.docx`

### セッション1 — Word 上で文書を作成（本文＋図形）

```text
あなたは MOS Word 365 演習アプリ向けの類題設計者です。
project5 に相当する演習文書を **1冊** 作成してください。
今回は **Word ファイルの作成とレイアウト仕様の出力** のみ。問題文は書きません。

【添付】Word問題文一覧_操作手順.csv（projectId=5 を参照）

【文書名】MOS_Word類題_project5_配置別_セット1_カフェ.docx
【テーマ】カフェ（バリスタ養成講座の案内文）

【レイアウト仕様 — 必ず遵守】
| 項目 | 値 |
|------|-----|
| セクション数 | 1 |
| 見出し構成 | バリスタ養成講座のご案内 / 講座概要 / 申込方法 |
| 余白 | 標準 |
| 印刷の向き | 縦 |
| スタイルセット | 基本（シンプル） |
| テーマカラー | 茶色 |
| 書式傾向 | 見出し下に薄茶背景 |
| 画像配置 | タイトル左（コーヒー豆）＋段落先頭（エスプレッソマシン）＋文末（カップ） |
| 改ページ位置 | なし |
| 印刷タイトル行 | なし（表なし） |
| 段落数（目安） | 20 |

【見出し構成】バリスタ養成講座のご案内 / 講座概要 / 申込方法
【段落数目安】20段落（2〜3ページ相当）

【必須構成とタスク用データ要件】

1. **5-1〜5-4 / 5-7 主画像**
   - 段落「6月10日より3日間の集中講座を開催します。」の先頭（画像「エスプレッソマシン」**未挿入**）
   - 5-2〜5-4・5-7の効果はすべて**未設定**

2. **5-5 / 5-6 副画像**
   - タイトル「バリスタ養成講座のご案内」の左側の画像「コーヒー豆」を**事前配置**（面取り未設定）
   - 5-6対象: タイトル「バリスタ養成講座のご案内」の左側の画像（説明・代替テキスト**未設定**）

3. **5-8 文末画像**
   - 文末に画像「カップ」を**事前配置**（背景削除のみは未実施）

【禁止】
- 教材の禁止文字列（5月21日より5日間、TOEICテスト対策セミナー、TOEIC、セミナー案内 等）
- 問題文の生成
- 受験者が行う操作を事前に完了させない

【出力】
1. docx ファイル
2. Markdown のレイアウト仕様書（見出し一覧・各タスクの段落先頭文・画像名）
```

### セッション2 — 精査

```text
あなたは MOS Word 類題の検証担当です。
添付の docx（MOS_Word類題_project5_配置別_セット1_カフェ.docx）を開き、以下を検証してください。問題文はまだ書きません。

【検証対象パラメータ（設計値）】
| taskId | パラメータ |
|--------|------------|
| 5-1 | anchorType=paragraphStart, targetParagraph=6月10日より3日間の集中講座を開催します。, imageName=エスプレッソマシン, wrapType=四角形, wrapTypeInternal=wdWrapSquare, primaryImageAnchor=paragraphStart |
| 5-2 | sameImageAs=task5_1, wrapType=貫通, wrapTypeInternal=wdWrapThrough |
| 5-3 | sameImageAs=task5_1, artEffect=鉛筆：スケッチ, artEffectInternal=artisticPencilSketch |
| 5-4 | sameImageAs=task5_1, softEdgePt=10, softEdgeLabel=ぼかし10ポイント |
| 5-5 | titleText=バリスタ養成講座のご案内, imageName=コーヒー豆, imageAnchor=titleLeft, bevelEffect=面取り 円, bevelInternal=bevelT+circle |
| 5-6 | sameImageAs=task5_5, altText=バリスタ講座案内, useAltTextWord=True |
| 5-7 | sameImageAs=task5_1, altTextDecorative=True, wording=装飾化 |
| 5-8 | imageName=カップ, imageAnchor=documentEnd, operation=removeBackgroundOnly |

【検証チェックリスト — PASS / FAIL】

| # | 項目 | 内容 |
|---|------|------|
| V1 | 5-1 | 挿入先・段落／見出しが設計どおり、主画像未挿入 |
| V2 | 5-5/6 | 副画像の位置・名称が設計どおり、効果・説明未設定 |
| V3 | 5-8 | 文末画像事前配置、背景処理未実施 |
| V4 | 画像名 | 3画像の名称が設計どおり |
| V5 | 禁止文字列 | TOEIC・5月21日関連なし |

【出力】精査結果サマリー + 修正済みパラメータ表（実文書の値で確定）
```

### セッション3 — 問題文生成

```text
【参照】完成初稿は task/MOS_Word類題_project5_配置別_5セット_問題文.md の
セット1（カフェ）と整合させること。

精査済みパラメータ表に基づき、タスク5-1〜5-8の問題文8件だけを出力してください。

【文体ルール — 操作種別は教材準拠、パラメータは設計値に従う】
- 5-1: 設計値の anchorType に応じて挿入（段落先頭／末尾／見出し下）。折り返しは「上下」以外。
- 5-2: 5-1と同一画像の折り返しを変更。「狭く」以外。
- 5-3: 5-1と同一画像にアート効果を設定。「水彩：スポンジ」以外。
- 5-4: 5-1と同一画像にぼかしを設定。「25ポイント」以外。
- 5-5: 設計値の imageAnchor（タイトル左／下／横、見出し右等）の画像に面取りを設定。「ハードエッジ」以外。
- 5-6: 設計値の位置の画像に説明を設定。useAltTextWord=false なら「説明」または「スクリーンリーダー用の説明」（「代替テキスト」という語は使わない）。
- 5-7: wording=装飾化 →「代替テキストを装飾化」／ wording=screenReaderHidden →「スクリーンリーダーに表示されないようにします」
- 5-8: operation=removeBackgroundOnly → 背景削除のみ／ addForegroundMark → 前景領域を一部追加でマーク＋背景以外削除しない

【出力形式】タスク5-1　...（8件）
```

### 設計値（taskParams）

| taskId | パラメータ |
|--------|------------|
| 5-1 | anchorType=paragraphStart, targetParagraph=6月10日より3日間の集中講座を開催します。, imageName=エスプレッソマシン, wrapType=四角形, wrapTypeInternal=wdWrapSquare, primaryImageAnchor=paragraphStart |
| 5-2 | sameImageAs=task5_1, wrapType=貫通, wrapTypeInternal=wdWrapThrough |
| 5-3 | sameImageAs=task5_1, artEffect=鉛筆：スケッチ, artEffectInternal=artisticPencilSketch |
| 5-4 | sameImageAs=task5_1, softEdgePt=10, softEdgeLabel=ぼかし10ポイント |
| 5-5 | titleText=バリスタ養成講座のご案内, imageName=コーヒー豆, imageAnchor=titleLeft, bevelEffect=面取り 円, bevelInternal=bevelT+circle |
| 5-6 | sameImageAs=task5_5, altText=バリスタ講座案内, useAltTextWord=True |
| 5-7 | sameImageAs=task5_1, altTextDecorative=True, wording=装飾化 |
| 5-8 | imageName=カップ, imageAnchor=documentEnd, operation=removeBackgroundOnly |

---

## §2 セット2 — 医療

**文書名:** `MOS_Word類題_project5_配置別_セット2_医療.docx`

### セッション1 — Word 上で文書を作成（本文＋図形）

```text
あなたは MOS Word 365 演習アプリ向けの類題設計者です。
project5 に相当する演習文書を **1冊** 作成してください。
今回は **Word ファイルの作成とレイアウト仕様の出力** のみ。問題文は書きません。

【添付】Word問題文一覧_操作手順.csv（projectId=5 を参照）

【文書名】MOS_Word類題_project5_配置別_セット2_医療.docx
【テーマ】医療（救急処置講習の案内文）

【レイアウト仕様 — 必ず遵守】
| 項目 | 値 |
|------|-----|
| セクション数 | 1 |
| 見出し構成 | 救急処置講習のご案内 / 講習内容 / 参加資格 |
| 余白 | やや狭い |
| 印刷の向き | 縦 |
| スタイルセット | 線（シンプル） |
| テーマカラー | 緑 |
| 書式傾向 | 本文段落に左罫線アクセント（緑） |
| 画像配置 | タイトル下（救急箱）＋段落末尾（聴診器）＋文末（救急車） |
| 改ページ位置 | なし |
| 印刷タイトル行 | なし（表なし） |
| 段落数（目安） | 22 |

【見出し構成】救急処置講習のご案内 / 講習内容 / 参加資格
【段落数目安】22段落（2〜3ページ相当）

【必須構成とタスク用データ要件】

1. **5-1〜5-4 / 5-7 主画像**
   - 段落「7月5日より2日間の実技講習を開催します。」の末尾（画像**未挿入**）
   - 5-2〜5-4・5-7の効果はすべて**未設定**

2. **5-5 / 5-6 副画像**
   - タイトル「救急処置講習のご案内」の下の画像「救急箱」を**事前配置**（面取り未設定）
   - 5-6対象: タイトル「救急処置講習のご案内」の下の画像（説明・代替テキスト**未設定**）

3. **5-8 文末画像**
   - 文末に画像「救急車」を**事前配置**（背景削除のみは未実施）

【禁止】
- 教材の禁止文字列（5月21日より5日間、TOEICテスト対策セミナー、TOEIC、セミナー案内 等）
- 問題文の生成
- 受験者が行う操作を事前に完了させない

【出力】
1. docx ファイル
2. Markdown のレイアウト仕様書（見出し一覧・各タスクの段落先頭文・画像名）
```

### セッション2 — 精査

```text
あなたは MOS Word 類題の検証担当です。
添付の docx（MOS_Word類題_project5_配置別_セット2_医療.docx）を開き、以下を検証してください。問題文はまだ書きません。

【検証対象パラメータ（設計値）】
| taskId | パラメータ |
|--------|------------|
| 5-1 | anchorType=paragraphEnd, targetParagraph=7月5日より2日間の実技講習を開催します。, imageName=聴診器, wrapType=貫通, wrapTypeInternal=wdWrapThrough, primaryImageAnchor=paragraphEnd |
| 5-2 | sameImageAs=task5_1, wrapType=四角形, wrapTypeInternal=wdWrapSquare |
| 5-3 | sameImageAs=task5_1, artEffect=マーカー, artEffectInternal=artisticMarker |
| 5-4 | sameImageAs=task5_1, softEdgePt=15, softEdgeLabel=ぼかし15ポイント |
| 5-5 | titleText=救急処置講習のご案内, imageName=救急箱, imageAnchor=underTitle, bevelEffect=面取り 角度, bevelInternal=bevelT+angle |
| 5-6 | sameImageAs=task5_5, altText=講習案内, useAltTextWord=True |
| 5-7 | sameImageAs=task5_1, altTextDecorative=True, wording=装飾化 |
| 5-8 | imageName=救急車, imageAnchor=documentEnd, operation=removeBackgroundOnly |

【検証チェックリスト — PASS / FAIL】

| # | 項目 | 内容 |
|---|------|------|
| V1 | 5-1 | 挿入先・段落／見出しが設計どおり、主画像未挿入 |
| V2 | 5-5/6 | 副画像の位置・名称が設計どおり、効果・説明未設定 |
| V3 | 5-8 | 文末画像事前配置、背景処理未実施 |
| V4 | 画像名 | 3画像の名称が設計どおり |
| V5 | 禁止文字列 | TOEIC・5月21日関連なし |

【出力】精査結果サマリー + 修正済みパラメータ表（実文書の値で確定）
```

### セッション3 — 問題文生成

```text
【参照】完成初稿は task/MOS_Word類題_project5_配置別_5セット_問題文.md の
セット2（医療）と整合させること。

精査済みパラメータ表に基づき、タスク5-1〜5-8の問題文8件だけを出力してください。

【文体ルール — 操作種別は教材準拠、パラメータは設計値に従う】
- 5-1: 設計値の anchorType に応じて挿入（段落先頭／末尾／見出し下）。折り返しは「上下」以外。
- 5-2: 5-1と同一画像の折り返しを変更。「狭く」以外。
- 5-3: 5-1と同一画像にアート効果を設定。「水彩：スポンジ」以外。
- 5-4: 5-1と同一画像にぼかしを設定。「25ポイント」以外。
- 5-5: 設計値の imageAnchor（タイトル左／下／横、見出し右等）の画像に面取りを設定。「ハードエッジ」以外。
- 5-6: 設計値の位置の画像に説明を設定。useAltTextWord=false なら「説明」または「スクリーンリーダー用の説明」（「代替テキスト」という語は使わない）。
- 5-7: wording=装飾化 →「代替テキストを装飾化」／ wording=screenReaderHidden →「スクリーンリーダーに表示されないようにします」
- 5-8: operation=removeBackgroundOnly → 背景削除のみ／ addForegroundMark → 前景領域を一部追加でマーク＋背景以外削除しない

【出力形式】タスク5-1　...（8件）
```

### 設計値（taskParams）

| taskId | パラメータ |
|--------|------------|
| 5-1 | anchorType=paragraphEnd, targetParagraph=7月5日より2日間の実技講習を開催します。, imageName=聴診器, wrapType=貫通, wrapTypeInternal=wdWrapThrough, primaryImageAnchor=paragraphEnd |
| 5-2 | sameImageAs=task5_1, wrapType=四角形, wrapTypeInternal=wdWrapSquare |
| 5-3 | sameImageAs=task5_1, artEffect=マーカー, artEffectInternal=artisticMarker |
| 5-4 | sameImageAs=task5_1, softEdgePt=15, softEdgeLabel=ぼかし15ポイント |
| 5-5 | titleText=救急処置講習のご案内, imageName=救急箱, imageAnchor=underTitle, bevelEffect=面取り 角度, bevelInternal=bevelT+angle |
| 5-6 | sameImageAs=task5_5, altText=講習案内, useAltTextWord=True |
| 5-7 | sameImageAs=task5_1, altTextDecorative=True, wording=装飾化 |
| 5-8 | imageName=救急車, imageAnchor=documentEnd, operation=removeBackgroundOnly |

---

## §3 セット3 — 製造

**文書名:** `MOS_Word類題_project5_配置別_セット3_製造.docx`

### セッション1 — Word 上で文書を作成（本文＋図形）

```text
あなたは MOS Word 365 演習アプリ向けの類題設計者です。
project5 に相当する演習文書を **1冊** 作成してください。
今回は **Word ファイルの作成とレイアウト仕様の出力** のみ。問題文は書きません。

【添付】Word問題文一覧_操作手順.csv（projectId=5 を参照）

【文書名】MOS_Word類題_project5_配置別_セット3_製造.docx
【テーマ】製造（品質検査研修の案内文）

【レイアウト仕様 — 必ず遵守】
| 項目 | 値 |
|------|-----|
| セクション数 | 1 |
| 見出し構成 | 品質検査研修のご案内 / 研修日程 / 受講対象 |
| 余白 | 広い |
| 印刷の向き | 縦 |
| スタイルセット | モダン |
| テーマカラー | 青灰 |
| 書式傾向 | 交互段落に薄灰背景縞 |
| 画像配置 | 見出し下（検査機器）＋タイトル右（安全帽）＋受講対象下（品質マーク）＋文末（工場） |
| 改ページ位置 | なし |
| 印刷タイトル行 | なし（表なし） |
| 段落数（目安） | 24 |

【見出し構成】品質検査研修のご案内 / 研修日程 / 受講対象
【段落数目安】24段落（2〜3ページ相当）

【必須構成とタスク用データ要件】

1. **5-1〜5-4 / 5-7 主画像**
   - 見出し「研修日程」の下（画像「検査機器」**未挿入**、折り返し狭くは未設定）
   - 5-2〜5-4・5-7の効果はすべて**未設定**

2. **5-5 / 5-6 副画像**
   - タイトル「品質検査研修のご案内」の右側の画像「安全帽」を**事前配置**（面取り未設定）
   - 5-6対象: 見出し「受講対象」の下の画像「品質マーク」を**事前配置**

3. **5-8 文末画像**
   - 文末に画像「工場」を**事前配置**（前景マーク追加＋背景削除は未実施）

【禁止】
- 教材の禁止文字列（5月21日より5日間、TOEICテスト対策セミナー、TOEIC、セミナー案内 等）
- 問題文の生成
- 受験者が行う操作を事前に完了させない

【出力】
1. docx ファイル
2. Markdown のレイアウト仕様書（見出し一覧・各タスクの段落先頭文・画像名）
```

### セッション2 — 精査

```text
あなたは MOS Word 類題の検証担当です。
添付の docx（MOS_Word類題_project5_配置別_セット3_製造.docx）を開き、以下を検証してください。問題文はまだ書きません。

【検証対象パラメータ（設計値）】
| taskId | パラメータ |
|--------|------------|
| 5-1 | anchorType=underHeading, anchorHeading=研修日程, targetParagraph=8月1日より4日間の研修を開催します。, imageName=検査機器, wrapType=狭く, wrapTypeInternal=wdWrapTight, primaryImageAnchor=underHeading |
| 5-2 | sameImageAs=task5_1, wrapType=背面, wrapTypeInternal=wdWrapBehind |
| 5-3 | sameImageAs=task5_1, artEffect=セピア, artEffectInternal=artisticSepia |
| 5-4 | sameImageAs=task5_1, softEdgePt=20, softEdgeLabel=ぼかし20ポイント |
| 5-5 | titleText=品質検査研修のご案内, imageName=安全帽, imageAnchor=titleRight, bevelEffect=面取り ソフトラウンド, bevelInternal=bevelT+softRound |
| 5-6 | anchorHeading=受講対象, imageName=品質マーク, imageAnchor=underHeading, altText=研修案内, useAltTextWord=False |
| 5-7 | sameImageAs=task5_1, altTextDecorative=True, wording=screenReaderHidden |
| 5-8 | imageName=工場, imageAnchor=documentEnd, operation=addForegroundMark |

【検証チェックリスト — PASS / FAIL】

| # | 項目 | 内容 |
|---|------|------|
| V1 | 5-1 | 挿入先・段落／見出しが設計どおり、主画像未挿入 |
| V2 | 5-5/6 | 副画像の位置・名称が設計どおり、効果・説明未設定 |
| V3 | 5-8 | 文末画像事前配置、背景処理未実施 |
| V4 | 画像名 | 3画像の名称が設計どおり |
| V5 | 禁止文字列 | TOEIC・5月21日関連なし |

【出力】精査結果サマリー + 修正済みパラメータ表（実文書の値で確定）
```

### セッション3 — 問題文生成

```text
【参照】完成初稿は task/MOS_Word類題_project5_配置別_5セット_問題文.md の
セット3（製造）と整合させること。

精査済みパラメータ表に基づき、タスク5-1〜5-8の問題文8件だけを出力してください。

【文体ルール — 操作種別は教材準拠、パラメータは設計値に従う】
- 5-1: 設計値の anchorType に応じて挿入（段落先頭／末尾／見出し下）。折り返しは「上下」以外。
- 5-2: 5-1と同一画像の折り返しを変更。「狭く」以外。
- 5-3: 5-1と同一画像にアート効果を設定。「水彩：スポンジ」以外。
- 5-4: 5-1と同一画像にぼかしを設定。「25ポイント」以外。
- 5-5: 設計値の imageAnchor（タイトル左／下／横、見出し右等）の画像に面取りを設定。「ハードエッジ」以外。
- 5-6: 設計値の位置の画像に説明を設定。useAltTextWord=false なら「説明」または「スクリーンリーダー用の説明」（「代替テキスト」という語は使わない）。
- 5-7: wording=装飾化 →「代替テキストを装飾化」／ wording=screenReaderHidden →「スクリーンリーダーに表示されないようにします」
- 5-8: operation=removeBackgroundOnly → 背景削除のみ／ addForegroundMark → 前景領域を一部追加でマーク＋背景以外削除しない

【出力形式】タスク5-1　...（8件）
```

### 設計値（taskParams）

| taskId | パラメータ |
|--------|------------|
| 5-1 | anchorType=underHeading, anchorHeading=研修日程, targetParagraph=8月1日より4日間の研修を開催します。, imageName=検査機器, wrapType=狭く, wrapTypeInternal=wdWrapTight, primaryImageAnchor=underHeading |
| 5-2 | sameImageAs=task5_1, wrapType=背面, wrapTypeInternal=wdWrapBehind |
| 5-3 | sameImageAs=task5_1, artEffect=セピア, artEffectInternal=artisticSepia |
| 5-4 | sameImageAs=task5_1, softEdgePt=20, softEdgeLabel=ぼかし20ポイント |
| 5-5 | titleText=品質検査研修のご案内, imageName=安全帽, imageAnchor=titleRight, bevelEffect=面取り ソフトラウンド, bevelInternal=bevelT+softRound |
| 5-6 | anchorHeading=受講対象, imageName=品質マーク, imageAnchor=underHeading, altText=研修案内, useAltTextWord=False |
| 5-7 | sameImageAs=task5_1, altTextDecorative=True, wording=screenReaderHidden |
| 5-8 | imageName=工場, imageAnchor=documentEnd, operation=addForegroundMark |

---

## §4 セット4 — 旅行

**文書名:** `MOS_Word類題_project5_配置別_セット4_旅行.docx`

### セッション1 — Word 上で文書を作成（本文＋図形）

```text
あなたは MOS Word 365 演習アプリ向けの類題設計者です。
project5 に相当する演習文書を **1冊** 作成してください。
今回は **Word ファイルの作成とレイアウト仕様の出力** のみ。問題文は書きません。

【添付】Word問題文一覧_操作手順.csv（projectId=5 を参照）

【文書名】MOS_Word類題_project5_配置別_セット4_旅行.docx
【テーマ】旅行（添乗員養成講座の案内文）

【レイアウト仕様 — 必ず遵守】
| 項目 | 値 |
|------|-----|
| セクション数 | 1 |
| 見出し構成 | 添乗員養成講座のご案内 / カリキュラム / 費用 |
| 余白 | 標準 |
| 印刷の向き | 縦 |
| スタイルセット | 伝統 |
| テーマカラー | 橙 |
| 書式傾向 | 見出し下二重罫線 |
| 画像配置 | 段落右（地球儀）＋費用横（スーツケース）＋文末（飛行機） |
| 改ページ位置 | なし |
| 印刷タイトル行 | なし（表なし） |
| 段落数（目安） | 21 |

【見出し構成】添乗員養成講座のご案内 / カリキュラム / 費用
【段落数目安】21段落（2〜3ページ相当）

【必須構成とタスク用データ要件】

1. **5-1〜5-4 / 5-7 主画像**
   - 見出し「カリキュラム」の下に段落「9月15日より5日間の講座を開催します。」を配置し、その右に画像「地球儀」**未挿入**
   - 5-2〜5-4・5-7の効果はすべて**未設定**

2. **5-5 / 5-6 副画像**
   - 見出し「費用」の右側の画像「スーツケース」を**事前配置**（面取り未設定）
   - 5-6対象: 見出し「費用」の右側の画像（説明・代替テキスト**未設定**）

3. **5-8 文末画像**
   - 文末に画像「飛行機」を**事前配置**（前景マーク追加＋背景削除は未実施）

【禁止】
- 教材の禁止文字列（5月21日より5日間、TOEICテスト対策セミナー、TOEIC、セミナー案内 等）
- 問題文の生成
- 受験者が行う操作を事前に完了させない

【出力】
1. docx ファイル
2. Markdown のレイアウト仕様書（見出し一覧・各タスクの段落先頭文・画像名）
```

### セッション2 — 精査

```text
あなたは MOS Word 類題の検証担当です。
添付の docx（MOS_Word類題_project5_配置別_セット4_旅行.docx）を開き、以下を検証してください。問題文はまだ書きません。

【検証対象パラメータ（設計値）】
| taskId | パラメータ |
|--------|------------|
| 5-1 | anchorType=besideParagraph, anchorHeading=カリキュラム, targetParagraph=9月15日より5日間の講座を開催します。, imageName=地球儀, wrapType=前面, wrapTypeInternal=wdWrapFront, primaryImageAnchor=besideParagraph |
| 5-2 | sameImageAs=task5_1, wrapType=上下, wrapTypeInternal=wdWrapTopBottom |
| 5-3 | sameImageAs=task5_1, artEffect=線画, artEffectInternal=artisticLineDrawing |
| 5-4 | sameImageAs=task5_1, softEdgePt=50, softEdgeLabel=ぼかし50ポイント |
| 5-5 | anchorHeading=費用, imageName=スーツケース, imageAnchor=headingRight, bevelEffect=面取り リラックスインセット, bevelInternal=bevelT+relaxedInset |
| 5-6 | sameImageAs=task5_1, altText=添乗員講座案内, useAltTextWord=False, altTextLabel=スクリーンリーダー用の説明 |
| 5-7 | sameImageAs=task5_1, altTextDecorative=True, wording=screenReaderHidden |
| 5-8 | imageName=飛行機, imageAnchor=documentEnd, operation=addForegroundMark |

【検証チェックリスト — PASS / FAIL】

| # | 項目 | 内容 |
|---|------|------|
| V1 | 5-1 | 挿入先・段落／見出しが設計どおり、主画像未挿入 |
| V2 | 5-5/6 | 副画像の位置・名称が設計どおり、効果・説明未設定 |
| V3 | 5-8 | 文末画像事前配置、背景処理未実施 |
| V4 | 画像名 | 3画像の名称が設計どおり |
| V5 | 禁止文字列 | TOEIC・5月21日関連なし |

【出力】精査結果サマリー + 修正済みパラメータ表（実文書の値で確定）
```

### セッション3 — 問題文生成

```text
【参照】完成初稿は task/MOS_Word類題_project5_配置別_5セット_問題文.md の
セット4（旅行）と整合させること。

精査済みパラメータ表に基づき、タスク5-1〜5-8の問題文8件だけを出力してください。

【文体ルール — 操作種別は教材準拠、パラメータは設計値に従う】
- 5-1: 設計値の anchorType に応じて挿入（段落先頭／末尾／見出し下）。折り返しは「上下」以外。
- 5-2: 5-1と同一画像の折り返しを変更。「狭く」以外。
- 5-3: 5-1と同一画像にアート効果を設定。「水彩：スポンジ」以外。
- 5-4: 5-1と同一画像にぼかしを設定。「25ポイント」以外。
- 5-5: 設計値の imageAnchor（タイトル左／下／横、見出し右等）の画像に面取りを設定。「ハードエッジ」以外。
- 5-6: 設計値の位置の画像に説明を設定。useAltTextWord=false なら「説明」または「スクリーンリーダー用の説明」（「代替テキスト」という語は使わない）。
- 5-7: wording=装飾化 →「代替テキストを装飾化」／ wording=screenReaderHidden →「スクリーンリーダーに表示されないようにします」
- 5-8: operation=removeBackgroundOnly → 背景削除のみ／ addForegroundMark → 前景領域を一部追加でマーク＋背景以外削除しない

【出力形式】タスク5-1　...（8件）
```

### 設計値（taskParams）

| taskId | パラメータ |
|--------|------------|
| 5-1 | anchorType=besideParagraph, anchorHeading=カリキュラム, targetParagraph=9月15日より5日間の講座を開催します。, imageName=地球儀, wrapType=前面, wrapTypeInternal=wdWrapFront, primaryImageAnchor=besideParagraph |
| 5-2 | sameImageAs=task5_1, wrapType=上下, wrapTypeInternal=wdWrapTopBottom |
| 5-3 | sameImageAs=task5_1, artEffect=線画, artEffectInternal=artisticLineDrawing |
| 5-4 | sameImageAs=task5_1, softEdgePt=50, softEdgeLabel=ぼかし50ポイント |
| 5-5 | anchorHeading=費用, imageName=スーツケース, imageAnchor=headingRight, bevelEffect=面取り リラックスインセット, bevelInternal=bevelT+relaxedInset |
| 5-6 | sameImageAs=task5_1, altText=添乗員講座案内, useAltTextWord=False, altTextLabel=スクリーンリーダー用の説明 |
| 5-7 | sameImageAs=task5_1, altTextDecorative=True, wording=screenReaderHidden |
| 5-8 | imageName=飛行機, imageAnchor=documentEnd, operation=addForegroundMark |

---

## §5 セット5 — 学習塾

**文書名:** `MOS_Word類題_project5_配置別_セット5_学習塾.docx`

### セッション1 — Word 上で文書を作成（本文＋図形）

```text
あなたは MOS Word 365 演習アプリ向けの類題設計者です。
project5 に相当する演習文書を **1冊** 作成してください。
今回は **Word ファイルの作成とレイアウト仕様の出力** のみ。問題文は書きません。

【添付】Word問題文一覧_操作手順.csv（projectId=5 を参照）

【文書名】MOS_Word類題_project5_配置別_セット5_学習塾.docx
【テーマ】学習塾（夏期集中講座の案内文）

【レイアウト仕様 — 必ず遵守】
| 項目 | 値 |
|------|-----|
| セクション数 | 1 |
| 見出し構成 | 夏期集中講座のご案内 / 時間割 / 料金 |
| 余白 | やや狭い |
| 印刷の向き | 縦 |
| スタイルセット | オフィス |
| テーマカラー | 紫 |
| 書式傾向 | 段落右余白広め |
| 画像配置 | タイトル下（タブレット）＋タイトル横（黒板）＋時間割下（学習アイコン）＋文末（鉛筆） |
| 改ページ位置 | なし |
| 印刷タイトル行 | なし（表なし） |
| 段落数（目安） | 19 |

【見出し構成】夏期集中講座のご案内 / 時間割 / 料金
【段落数目安】19段落（2〜3ページ相当）

【必須構成とタスク用データ要件】

1. **5-1〜5-4 / 5-7 主画像**
   - タイトル「夏期集中講座のご案内」の下（画像「タブレット」**未挿入**）
   - 5-2〜5-4・5-7の効果はすべて**未設定**

2. **5-5 / 5-6 副画像**
   - タイトル「夏期集中講座のご案内」の横の画像「黒板」を**事前配置**（面取り未設定）
   - 5-6対象: 見出し「時間割」の下の画像「学習アイコン」を**事前配置**

3. **5-8 文末画像**
   - 文末に画像「鉛筆」を**事前配置**（前景マーク追加＋背景削除は未実施）

【禁止】
- 教材の禁止文字列（5月21日より5日間、TOEICテスト対策セミナー、TOEIC、セミナー案内 等）
- 問題文の生成
- 受験者が行う操作を事前に完了させない

【出力】
1. docx ファイル
2. Markdown のレイアウト仕様書（見出し一覧・各タスクの段落先頭文・画像名）
```

### セッション2 — 精査

```text
あなたは MOS Word 類題の検証担当です。
添付の docx（MOS_Word類題_project5_配置別_セット5_学習塾.docx）を開き、以下を検証してください。問題文はまだ書きません。

【検証対象パラメータ（設計値）】
| taskId | パラメータ |
|--------|------------|
| 5-1 | anchorType=underTitle, titleText=夏期集中講座のご案内, targetParagraph=7月20日より10日間の講座を開催します。, imageName=タブレット, wrapType=背面, wrapTypeInternal=wdWrapBehind, primaryImageAnchor=underTitle |
| 5-2 | sameImageAs=task5_1, wrapType=前面, wrapTypeInternal=wdWrapFront |
| 5-3 | sameImageAs=task5_1, artEffect=チョーク, artEffectInternal=artisticChalk |
| 5-4 | sameImageAs=task5_1, softEdgePt=35, softEdgeLabel=ぼかし35ポイント |
| 5-5 | titleText=夏期集中講座のご案内, imageName=黒板, imageAnchor=besideTitle, bevelEffect=面取り クールスラント, bevelInternal=bevelT+coolSlant |
| 5-6 | anchorHeading=時間割, imageName=学習アイコン, imageAnchor=underHeading, altText=夏期講座案内, useAltTextWord=False |
| 5-7 | sameImageAs=task5_1, altTextDecorative=True, wording=screenReaderHidden |
| 5-8 | imageName=鉛筆, imageAnchor=documentEnd, operation=addForegroundMark |

【検証チェックリスト — PASS / FAIL】

| # | 項目 | 内容 |
|---|------|------|
| V1 | 5-1 | 挿入先・段落／見出しが設計どおり、主画像未挿入 |
| V2 | 5-5/6 | 副画像の位置・名称が設計どおり、効果・説明未設定 |
| V3 | 5-8 | 文末画像事前配置、背景処理未実施 |
| V4 | 画像名 | 3画像の名称が設計どおり |
| V5 | 禁止文字列 | TOEIC・5月21日関連なし |

【出力】精査結果サマリー + 修正済みパラメータ表（実文書の値で確定）
```

### セッション3 — 問題文生成

```text
【参照】完成初稿は task/MOS_Word類題_project5_配置別_5セット_問題文.md の
セット5（学習塾）と整合させること。

精査済みパラメータ表に基づき、タスク5-1〜5-8の問題文8件だけを出力してください。

【文体ルール — 操作種別は教材準拠、パラメータは設計値に従う】
- 5-1: 設計値の anchorType に応じて挿入（段落先頭／末尾／見出し下）。折り返しは「上下」以外。
- 5-2: 5-1と同一画像の折り返しを変更。「狭く」以外。
- 5-3: 5-1と同一画像にアート効果を設定。「水彩：スポンジ」以外。
- 5-4: 5-1と同一画像にぼかしを設定。「25ポイント」以外。
- 5-5: 設計値の imageAnchor（タイトル左／下／横、見出し右等）の画像に面取りを設定。「ハードエッジ」以外。
- 5-6: 設計値の位置の画像に説明を設定。useAltTextWord=false なら「説明」または「スクリーンリーダー用の説明」（「代替テキスト」という語は使わない）。
- 5-7: wording=装飾化 →「代替テキストを装飾化」／ wording=screenReaderHidden →「スクリーンリーダーに表示されないようにします」
- 5-8: operation=removeBackgroundOnly → 背景削除のみ／ addForegroundMark → 前景領域を一部追加でマーク＋背景以外削除しない

【出力形式】タスク5-1　...（8件）
```

### 設計値（taskParams）

| taskId | パラメータ |
|--------|------------|
| 5-1 | anchorType=underTitle, titleText=夏期集中講座のご案内, targetParagraph=7月20日より10日間の講座を開催します。, imageName=タブレット, wrapType=背面, wrapTypeInternal=wdWrapBehind, primaryImageAnchor=underTitle |
| 5-2 | sameImageAs=task5_1, wrapType=前面, wrapTypeInternal=wdWrapFront |
| 5-3 | sameImageAs=task5_1, artEffect=チョーク, artEffectInternal=artisticChalk |
| 5-4 | sameImageAs=task5_1, softEdgePt=35, softEdgeLabel=ぼかし35ポイント |
| 5-5 | titleText=夏期集中講座のご案内, imageName=黒板, imageAnchor=besideTitle, bevelEffect=面取り クールスラント, bevelInternal=bevelT+coolSlant |
| 5-6 | anchorHeading=時間割, imageName=学習アイコン, imageAnchor=underHeading, altText=夏期講座案内, useAltTextWord=False |
| 5-7 | sameImageAs=task5_1, altTextDecorative=True, wording=screenReaderHidden |
| 5-8 | imageName=鉛筆, imageAnchor=documentEnd, operation=addForegroundMark |

---

## 完成問題文（初稿）

### MOS_Word類題_project5_配置別_セット1_カフェ.docx

あなたはバリスタ養成講座の案内文を作成しています。  
問題）  
タスク5-1　「6月10日より3日間の集中講座を開催します。」の段落の先頭に、画像「エスプレッソマシン」を挿入します。文字列の折り返しは「四角形」にします。
タスク5-2　「6月10日より3日間の集中講座を開催します。」の段落の先頭の画像の文字列の折り返しを「貫通」に変更します。
タスク5-3　「6月10日より3日間の集中講座を開催します。」の段落の先頭の画像にアート効果「鉛筆：スケッチ」を設定します。
タスク5-4　「6月10日より3日間の集中講座を開催します。」の段落の先頭の画像に図の効果「ぼかし10ポイント」を設定します。
タスク5-5　「バリスタ養成講座のご案内」のタイトルの左側の画像に、図の効果「面取り 円」を設定します。
タスク5-6　「バリスタ養成講座のご案内」のタイトルの左側の画像がスクリーンリーダーに表示されるように、代替テキスト「"バリスタ講座案内"」を設定します。
タスク5-7　「6月10日より3日間の集中講座を開催します。」の段落の先頭の画像の代替テキストを装飾化します。
タスク5-8　文末の画像の背景を削除します。背景以外は削除しないようにします。

### MOS_Word類題_project5_配置別_セット2_医療.docx

あなたは救急処置講習の案内文を作成しています。  
問題）  
タスク5-1　「7月5日より2日間の実技講習を開催します。」の段落の末尾に、画像「聴診器」を挿入します。文字列の折り返しは「貫通」にします。
タスク5-2　「7月5日より2日間の実技講習を開催します。」の段落の末尾の画像の文字列の折り返しを「四角形」に変更します。
タスク5-3　「7月5日より2日間の実技講習を開催します。」の段落の末尾の画像にアート効果「マーカー」を設定します。
タスク5-4　「7月5日より2日間の実技講習を開催します。」の段落の末尾の画像に図の効果「ぼかし15ポイント」を設定します。
タスク5-5　「救急処置講習のご案内」のタイトルの下の画像に、図の効果「面取り 角度」を設定します。
タスク5-6　「救急処置講習のご案内」のタイトルの下の画像がスクリーンリーダーに表示されるように、代替テキスト「"講習案内"」を設定します。
タスク5-7　「7月5日より2日間の実技講習を開催します。」の段落の末尾の画像の代替テキストを装飾化します。
タスク5-8　文末の画像の背景を削除します。背景以外は削除しないようにします。

### MOS_Word類題_project5_配置別_セット3_製造.docx

あなたは品質検査研修の案内文を作成しています。  
問題）  
タスク5-1　見出し「研修日程」の下に、画像「検査機器」を挿入します。文字列の折り返しは「狭く」にします。
タスク5-2　見出し「研修日程」の下の画像の文字列の折り返しを「背面」に変更します。
タスク5-3　見出し「研修日程」の下の画像にアート効果「セピア」を設定します。
タスク5-4　見出し「研修日程」の下の画像に図の効果「ぼかし20ポイント」を設定します。
タスク5-5　「品質検査研修のご案内」のタイトルの右側の画像に、図の効果「面取り ソフトラウンド」を設定します。
タスク5-6　見出し「受講対象」の下の画像がスクリーンリーダーに表示されるように、説明「"研修案内"」を設定します。
タスク5-7　見出し「研修日程」の下の画像がスクリーンリーダーに表示されないようにします。
タスク5-8　文末の画像で、前景として残す領域を一部追加でマークします。背景以外は削除しないようにします。

### MOS_Word類題_project5_配置別_セット4_旅行.docx

あなたは添乗員養成講座の案内文を作成しています。  
問題）  
タスク5-1　見出し「カリキュラム」の下の段落「9月15日より5日間の講座を開催します。」の右に、画像「地球儀」を挿入します。文字列の折り返しは「前面」にします。
タスク5-2　見出し「カリキュラム」の下の段落「9月15日より5日間の講座を開催します。」の右の画像の文字列の折り返しを「上下」に変更します。
タスク5-3　見出し「カリキュラム」の下の段落「9月15日より5日間の講座を開催します。」の右の画像にアート効果「線画」を設定します。
タスク5-4　見出し「カリキュラム」の下の段落「9月15日より5日間の講座を開催します。」の右の画像に図の効果「ぼかし50ポイント」を設定します。
タスク5-5　見出し「費用」の右側の画像に、図の効果「面取り リラックスインセット」を設定します。
タスク5-6　見出し「カリキュラム」の下の段落「9月15日より5日間の講座を開催します。」の右の画像がスクリーンリーダーに表示されるように、スクリーンリーダー用の説明「"添乗員講座案内"」を設定します。
タスク5-7　見出し「カリキュラム」の下の段落「9月15日より5日間の講座を開催します。」の右の画像がスクリーンリーダーに表示されないようにします。
タスク5-8　文末の画像で、前景として残す領域を一部追加でマークします。背景以外は削除しないようにします。

### MOS_Word類題_project5_配置別_セット5_学習塾.docx

あなたは夏期集中講座の案内文を作成しています。  
問題）  
タスク5-1　「夏期集中講座のご案内」のタイトルの下に、画像「タブレット」を挿入します。文字列の折り返しは「背面」にします。
タスク5-2　「夏期集中講座のご案内」のタイトルの下の画像の文字列の折り返しを「前面」に変更します。
タスク5-3　「夏期集中講座のご案内」のタイトルの下の画像にアート効果「チョーク」を設定します。
タスク5-4　「夏期集中講座のご案内」のタイトルの下の画像に図の効果「ぼかし35ポイント」を設定します。
タスク5-5　「夏期集中講座のご案内」のタイトルの横の画像に、図の効果「面取り クールスラント」を設定します。
タスク5-6　見出し「時間割」の下の画像がスクリーンリーダーに表示されるように、説明「"夏期講座案内"」を設定します。
タスク5-7　「夏期集中講座のご案内」のタイトルの下の画像がスクリーンリーダーに表示されないようにします。
タスク5-8　文末の画像で、前景として残す領域を一部追加でマークします。背景以外は削除しないようにします。

## 改訂履歴

| 日付 | 内容 |
|------|------|
| 2026-06-26 | Word Project5 配置別5セット初版 |
