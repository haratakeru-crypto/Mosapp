# New_MOSWordVSTOAddIn — ログに残るコマンド一覧と操作対応

このドキュメントは、`New_MOSWordVSTOAddIn` が `Logger.LogCommand` で `mos_word_log.txt` に記録する **コマンドID** と、ユーザーが行う **Word 上の操作（相当するボタン・機能）** の対応を整理したものです。

## ログの形式

ログファイルの既定パス: `%TEMP%\mos_word_log.txt`（`Logger.GetLogFilePath()` と一致）

### 行種別（破壊的操作検知）

| 種別 | 例 | 記録主体 |
|------|-----|----------|
| TaskStart | `[ts] [TaskStart] 7-3-0` | 試験アプリ（`LogReader.LogTaskStart`） |
| Operation | `[ts] [Task 7-3-0] [Op] RibbonCommand Cut` | VSTO（`Logger.LogOperation`） |
| 正解証跡 | `[ts] [Project7] [Task7-3] [IntegralHeader] Executed` | VSTO（`Logger.LogTaskEvidence`） |
| デバッグ | `[ts] [コマンドID] Executed` | VSTO（`Logger.LogCommand`、採点・破壊判定では無視） |

従来の証跡行の例:

```text
[2025-03-27 12:34:56] [Project7] [Task7-3] [IntegralHeader] Executed
```

---

## 1. リボンでフックした組み込みコマンド（`Ribbon.xml` の `<commands>`）

[`New_MOSWordVSTOAddIn/New_MOSWordVSTOAddIn/Ribbon.xml`](../../New_MOSWordVSTOAddIn/New_MOSWordVSTOAddIn/Ribbon.xml) の `<commands>` に列挙した idMso が `CommandOnAction` に接続されています。以下は代表例です。**ログに書かれるID** は、正規化ルールにより **元の idMso と異なる場合** があります。

| ログに記録されるコマンドID | Word の操作（目安） | 元の idMso / 備考 |
|---------------------------|---------------------|---------------------|
| `Cut` | **ホーム** の **切り取り** など | `Cut` |
| `Copy` | **コピー** | `Copy` |
| `Paste` | **貼り付け** | `Paste` |
| `FontColorMoreColorsDialog` | **フォントの色** → **その他の色** | 色ギャラリー本体は `<commands>` でフック不可のため未接続。ダイアログ経由のみログ |
| `ShowAll` | **編集記号の表示/非表示** | Word の `<commands>` では **`ShowAll` が不明 ID** のため **XML 未接続**。**ThisAddIn** の `View.ShowAll` / `Options.ShowAll` ポーリングのみ |
| `PageMarginsModerate` | **余白** の **やや狭い** など | Word の `<commands>` では **`PageMarginsModerate` が不明 ID** のため **XML 未接続**。採点は文書状態（余白値）を主判定 |
| `PageOrientationPortraitLandscape` | **レイアウト** の **向き** | `PageOrientationPortraitLandscape`。縦横個別 idMso は XML 未接続（**ThisAddIn** の全セクション向きポーリングで補完） |
| `ColumnBreak` | **レイアウト** → **区切り** → **段区切り** | リボン idMso は環境で無効のため **XML 未接続**。**ThisAddIn** の段区切り文字数ポーリングのみ |
| `ColumnsLeft`（互換） | **レイアウト** → **段組み** → **その他の段組み** | `ColumnsDialog` を **`ColumnsLeft`** に正規化（`ColumnsLeft`/`ColumnsRight` idMso は Word で無効） |
| `InsertSectionBreakNextPage` | **次のページから開始** 等 | `SectionBreakInsert` のみフックし **`InsertSectionBreakNextPage`** に正規化（`InsertSectionBreakNextPage` idMso は XML 未接続） |
| `PageBorders` | **デザイン** の **ページの罫線**／**枠線と網掛け**（ページ罫線タブ） | Word の XML では **`PageBorderAndShadingDialog`** のみ有効。`PageBorders` / `PageBorderOptionsDialog` は不明 ID のため未使用。ログは **`PageBorders`** に正規化 |
| `StyleSetLineSimple` | **デザイン** の **スタイルセット**（線・シンプル等） | Word の `<commands>` では **`StyleSetGallery` が不明 ID** のため **XML 未接続**。**ThisAddIn** が見出し1の下罫線（単線0.5pt）へ変化したときにログ |
| `TableConvertTextToTable` | **表に変換** | `ConvertTextToTable` を **`TableConvertTextToTable`** に正規化（`Ribbon.cs`） |
| `FileSaveAs` | **ファイル** → **名前を付けて保存**（汎用） | **Ribbon `idMso="FileSaveAs"`**（採点 7-4/7-5 では未使用） |
| `FileSaveAsTxt` | **7-4**: 書式なし **`朗読会.txt`** として保存 | **ポーリング**／**`DocumentBeforeSave`** で **`朗読会.txt` へ遷移**したとき（事前に txt があっても、今回の操作のみ） |
| `FileSaveAsDocm` | **7-5**: **`朗読会.docm`** として保存 | **ポーリング**／**`DocumentBeforeSave`** で **`朗読会.docm` へ遷移**したとき（7-4 の txt ログとは別 ID） |
| `IntegralHeader` | **挿入** → **ヘッダー** → **インテグラル** | **ポーリング**で `WordChecker1_7` と同一判定が false→true のとき記録。**7-3** は XML OR このログ |
| `UpgradeDocument` | **ファイル** → **情報** → **変換**（互換モード解除） | **Ribbon の `idMso` では onAction が発火しない**（実機確認済み）。**7-1 用のログは `ThisAddIn` ポーリング**で、`.doc` かつ `CompatibilityMode` が **非 2013 → Word 2013（15）** に遷移したときに **`UpgradeDocument` として記録**する。採点は `WordChecker1_7` で **現在状態 OR このログ**（5-1 と同型） |

### コード上の正規化（`Ribbon.cs`）

- `ConvertTextToTable` → `TableConvertTextToTable`
- `SectionBreakInsert` → `InsertSectionBreakNextPage`
- `ColumnsDialog` → `ColumnsLeft`（チェッカーが `ColumnsLeft` / `ColumnsRight` のいずれかを参照するため）
- `PageBorderAndShadingDialog` / `PageBorders` / `PageBorderOptionsDialog` / （コールバックで受けた場合）→ `PageBorders`
- （`StyleSetGallery` は Word で Ribbon XML 不可）`StyleSetLineSimple` は **ポーリング**で記録
- 上記以外は `control.Id` がそのままログID

---

## 2. タイマーによる検知（`ThisAddIn` のポーリング）

idMso で拾えない操作や、文書状態の変化から補完的に記録します（**間隔 1200ms**。重い判定は **5 ティックに 1 回**＝約 6 秒ごと）。

| ログに記録されるコマンドID | 検知内容（ユーザー操作の目安） |
|---------------------------|--------------------------------|
| `ShowAll` | **編集記号** の切り替え（`View.ShowAll` / `Options.ShowAll` の変化） |
| `PageOrientationPortraitLandscape` | **全セクション**の向きフィンガープリント（各 `PageSetup.Orientation` を連結）が変化したとき。先頭セクションだけでは 3-3 と採点チェッカーが一致しないため、**セクション2のみ横向き**なども検知する |
| `PageBorders` | 先頭セクションの **ページ上辺・下辺罫線**のスナップショットが変化したとき |
| `StyleSetLineSimple` | **見出し 1** の **下罫線**がチェッカー（`WordChecker1_4`）と同条件の **単線 0.5pt** になった遷移 |
| `ColumnBreak` | 文書内の **段区切り**（文字コード 14）の **個数が増えた** とき（重いチェック実行時のみ） |
| `UpgradeDocument` | **7-1**: `.doc` で **互換モードが非 Word2013 → Word2013（15）** に変化したとき（**変換** 操作の補完ログ。Ribbon の idMso は発火しない） |
| `IntegralHeader` | **7-3**: `EvaluateIntegralHeaderPresenceForPolling` が **false→true**（約 6 秒間隔のヘビーポーリング） |
| `FileSaveAsTxt` | **7-4**: ActiveDocument が **`朗読会.txt` へ遷移**（初回オープン時はベースラインのみ・ログなし） |
| `FileSaveAsDocm` | **7-5**: ActiveDocument が **`朗読会.docm` へ遷移**（同上） |

**二重記録の抑止:** リボンで既に `PageOrientationPortraitLandscape` / `PageBorders` を記録した直後は、同一操作に対するポーリング側のログを数ティック抑止する（`RegisterRibbonLogged*`）。`StyleSetLineSimple` はリボン未フックのためポーリングのみ。

**注意:** `ColumnBreak` は、リボンとポーリングの **両方** から同じIDが記録される場合があります。

---

## 3. 文書切替時のベースライン（`Application.DocumentChange`）

別文書を開いた直後や、ホストアプリの **リセットでログファイルを削除**した直後でも、VSTO 側のメモリ上の「前回値」は自動では消えません。`DocumentChange` で **現在の文書**から向き・罫線・見出し1・段区切り・編集記号・**7-1 用互換モード**の **ベースラインを再取得**し（ログは出さない）、誤検知やリセット直後のゴーストログを防ぎます。

ホストの `LogReader.ClearLog()` はファイルだけ削除するため、**文書を開き直す**か **編集操作で状態が変わる**まで、ポーリングは新しいベースラインと比較します。

---

## 4. `InsertSectionBreakNextPage` とポーリングについて

**次のページから開始** の操作は **`Ribbon.xml` の `SectionBreakInsert` フック**（ログ ID は `InsertSectionBreakNextPage` に正規化）で記録します。セクション数のポーリングによる `InsertSectionBreakNextPage` は **実装していません**（リボンとの二重記録・リセット後の誤記録を避けるため）。

---

## 5. 実装はあるが XML 未接続のハンドラ（現状ログに出にくい）

| 想定されるログ | 内容 | 備考 |
|----------------|------|------|
| `TocAutomatic2` | 目次ギャラリーで特定スタイルを選んだとき | `TableOfContentsGalleryOnAction` 内。現行 `Ribbon.xml` に **ギャラリーのフックなし** のため、通常は呼ばれない |
| `TableOfContentsGallery-{selectedId}` | 上記以外の目次タイプ選択時 | 同上 |

---

## 6. 「MOSデバッグ」タブのカスタムボタン（ログには **残らない**）

| ボタンラベル | 処理 |
|-------------|------|
| **ログパス表示** | ログファイルのパスをメッセージ表示（`Logger.LogCommand` なし） |
| **ログ内容表示** | ログファイルの内容をフォーム表示（`Logger.LogCommand` なし） |

---

## 参照ソース（プロジェクト内）

| ファイル | 役割 |
|----------|------|
| `New_MOSWordVSTOAddIn/Logger.cs` | `LogCommand`、ログファイル名・パス |
| `New_MOSWordVSTOAddIn/Ribbon.xml` | `<commands>` の idMso と MOSデバッグタブ |
| `New_MOSWordVSTOAddIn/Ribbon.cs` | `CommandOnAction` の正規化、`TableOfContentsGalleryOnAction` |
| `New_MOSWordVSTOAddIn/ThisAddIn.cs` | `ShowAll` / 全セクション向き・罫線・スタイルセットのポーリング、**7-1 互換モード遷移（`UpgradeDocument` ログ）**、`DocumentChange` でのベースライン更新、段区切り数のポーリング |

---

*最終更新: ソース `New_MOSWordVSTOAddIn`（Ribbon `<commands>` 復元・向きポーリング全セクション化・次ページ区切りポーリングなし）に基づく。*
