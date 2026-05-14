# プロジェクト7 採点機能 改修中ドキュメント

プロジェクト7（互換モード解除／文書プロパティ／ヘッダー挿入／別名保存）全5タスクの採点ロジック改修方針・進捗をまとめる。
本ドキュメントは改修の **作業中ドキュメント** であり、各タスクの方針が確定し実装が完了次第、随時追記・修正していく。
すべてのタスクが完了したら `WordChecker1_7_完了レポート.md` にリネーム／再構成する想定。

> 設計指針・判定方針の決定フロー・実装上の注意点は、[`WordChecker1_6_完了レポート.md`](./WordChecker1_6_完了レポート.md) で体系化された知見を継承する。本ドキュメントでは **プロジェクト7固有の方針・判定ロジック・落とし穴のみ** を追記する。

---

## 現状サマリ

| タスク | 改修ステータス | 採点結果 |
| :--- | :--- | :--- |
| 7-1 | **実装完了**（プランE: State-only Strict） | ○ |
| 7-2 | **実装完了**（COM厳密一致 + `.Value` バグ修正） | ○ |
| 7-3 | **ダンプ採取フェーズ**（識別タグ調査中、判定方針=「中」XMLのみ） | × |
| 7-4 | 未着手 | × |
| 7-5 | 未着手 | × |

改修順序は **7-1 → 7-2 → 7-3 → 7-4 → 7-5** で進める。

---

## プロジェクト7 タスク一覧（CSV準拠）

| タスク | 問題文 | 想定操作（CSVより） |
| :--- | :--- | :--- |
| 7-1 | 文書の互換モードを解除します。メッセージが表示された場合は「OK」をクリックします。 | ［ファイル］→［情報］→［変換］→［OK］ |
| 7-2 | 文章のプロパティの会社名に「ラビット出版」と設定します。 | ［ファイル］→［情報］→［プロパティをすべて表示］→［会社］に入力 |
| 7-3 | 文書に「インテグラル」のヘッダーを挿入します。 | ［挿入］→［ヘッダー］→［インテグラル］ |
| 7-4 | 文書に「朗読会」という名前を付けてテキストファイルとして保存します。ファイルの変換は既定値のままにします。 | ［ファイル］→［名前を付けて保存］→［参照］→ファイル名「朗読会」／種類「書式なし(*.txt)」 |
| 7-5 | この文書のコピーをマクロ有効文書として保存します。「名前を付けて保存」画面で読み取りパスワードを「abc」に設定すること。 | ［ファイル］→［名前を付けて保存］→［参照］→種類「マクロ有効文書(*.docm)」→［ツール］→［全般オプション］→読み取りパスワード「abc」 |

---

## 既存実装の問題点（改修の動機）

`MOSapp/MOS Word app/Libraries/WordChecker/Group1/WordChecker1_7.cs` の現状：

| タスク | 現状 | 問題点 | 重大度 |
| :--- | :--- | :--- | :--- |
| 7-1 | ~~`fileStateCheck = true` 固定 + `UpgradeDocument` のログAND~~ → **改修済み**（State-only Strict） | ~~ファイル状態の検証が完全に未実装。ログだけ存在すれば常に合格になる擬似実装。~~ → 実装完了 | ~~高~~ → 解消 |
| 7-2 | ~~`BuiltInDocumentProperties["Company"]` + `Contains("ラビット出版")`~~ → **改修済み**（完全一致判定） | ~~`Contains` による部分一致のため、「ラビット出版社」「ラビット出版部」「○○ラビット出版」も合格してしまう。MOSの採点ポリシー（完全一致）から逸脱。~~ → 実装完了 | ~~中~~ → 解消 |
| 7-3 | `Sections[1].Headers[wdHeaderFooterPrimary].Range.Text.Contains("インテグラル")` | 「インテグラル」ヘッダー（ビルトイン Building Block）は意匠的デザインのみで、**「インテグラル」というテキスト文字列はヘッダー内に存在しない**。よって原理的に常に false。 | 高 |
| 7-4 | `fileStateCheck = true` 固定 + `FileSaveAs` のログAND | 保存ファイル名「朗読会」・形式「.txt」が未検証。ログだけ存在すれば常に合格。 | 高 |
| 7-5 | `fileStateCheck = true` 固定 + `FileSaveAs` のログAND | 保存形式「.docm」・読み取りパスワード「abc」設定が未検証。ログだけ存在すれば常に合格。 | 高 |

→ **5タスク中4タスクで判定ロジックの実体化が必要**、1タスク（7-2）で判定厳密化が必要。
→ 「現状すべて × となっている」原因は **ログ自体がそもそも出ていない可能性** が高い。Ribbon.xml のフック状況を要確認。

---

## 共通の設計指針（プロジェクト6から継承）

1. **ターゲット特定ロジック**：見出し近傍の表は `GetTableXmlNearText` で特定する。ただしプロジェクト7は表操作が無いため本指針の出番は限定的。
2. **アドイン（VSTO）の制約**：`idMso` でフックできないコマンドが存在する。ログAND方針を採用する前に hookability を確認する。
3. **COM最小化**：`document.WordOpenXML` または `BuiltInDocumentProperties` で完結できるなら COM/XML で完結させる。
4. **既存実装の `Contains` パターンは要注意**：MOS採点は基本的に「完全一致」または「特定の構造特徴」で判定するべき。プロジェクト6 の 6-3 で確立した「誤学習を防ぐ厳密判定」の思想を継承する。

---

## 判定方針の決定フロー（プロジェクト6から継承）

```
[1] 操作の最終状態に、ユニークなXMLタグや属性値が残るか？
    ├─ 残る   → [2] へ
    └─ 残らない → COMのまま、または操作ログのみで判定（VSTO制約に注意）

[2] 同じXML最終状態を「誤った操作・手動操作」で作れてしまうか？
    ├─ 作れる   → XML判定 + ログ記録 AND（誤学習防止）
    └─ 作れない → XML判定のみ（リボン/ダイアログ/右クリック等の正解ルートをすべて許容）

[3] 既に正しく動作していて、対象が事実上ユニークに絞れるなら？
    └─ 現状維持も有効な選択肢（改修リスク回避）
```

### プロジェクト7 特有の判定軸

プロジェクト7は **「ファイル状態（保存形式・プロパティ）に対する操作」** が中心であり、プロジェクト6（XML上の構造変化）とは判定対象の性質が異なる。

| 操作の性質 | 判定対象（候補） | 想定方針 |
| :--- | :--- | :--- |
| **互換モード解除** | 拡張子`.doc`（変わらない）＋`Document.CompatibilityMode == 15`（COM） | **State-only Strict**（`.doc + 15` がユニークシグネチャ。`UpgradeDocument` はフック不能） |
| **文書プロパティ変更** | `BuiltInDocumentProperties["Company"]` の **完全一致**（COM） | **COM判定のみ・厳密一致**（空白/全角半角の差を区別） |
| **ヘッダー挿入** | `header*.xml` の **Building Block 識別子**（`<w:docPart w:val="..."/>` または SDT属性内の英語キーワード） | **XML判定のみ**（要実機ダンプ） |
| **別名保存（ファイル形式）** | **保存先ディレクトリ走査**（`File.Exists`）＋拡張子＋（必要なら）形式の妥当性 | **ファイルシステム + ログAND**（複合判定） |
| **パスワード保護** | 保存ファイルが ZIP として開けないか／暗号化されているか | **ファイルシステム + ログAND**（複合判定） |

---

## タスク別 最終実装サマリ（進捗管理表）

| タスク | 内容 | 現状 | 推奨方針 | 判定ポイント（仮） | 実装状況 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **7-1** | 互換モードの解除 | ○（実装完了） | **State-only Strict** | 拡張子`.doc` + `CompatibilityMode == 15`（ユニークシグネチャ） | **完了** |
| **7-2** | 文書プロパティ：会社名 | ○（実装完了） | **COM厳密一致 + `.Value`経由** | `BuiltInDocumentProperties["Company"].Value == "ラビット出版"`（完全一致、空白/全角半角の差を区別） | **完了** |
| **7-3** | 「インテグラル」ヘッダー挿入 | × | XMLのみ | `word/header*.xml` 内の `<w:docPart w:val="..."/>` または SDT属性に Building Block 識別子 | 未着手（要XMLダンプ） |
| **7-4** | テキストファイル「朗読会」として保存 | × | ファイル存在 + ログAND | 元ファイル同一フォルダに `朗読会.txt` が存在し、`FileSaveAs` ログあり | 未着手 |
| **7-5** | マクロ有効文書として保存＋パスワード「abc」 | × | ファイル存在 + 暗号化判定 + ログAND | 元ファイル同一フォルダに同名 `.docm` が存在、ZIP として開けない（暗号化）、`FileSaveAs` ログあり | 未着手 |

---

## 各タスクの方針詳細

### 7-1 互換モードの解除

- **ステータス**: **実装完了（プランE採用）**
- **判定対象（確定）**:
  - **拡張子**: `.doc`（変換直後の未保存状態。詳細は下記の「重要発見」参照）
  - **COM: `Document.CompatibilityMode`**（`WdCompatibilityMode.wdWord2013` = 15 であれば最新形式）
  - **操作ログ**: **使用しない**（`UpgradeDocument` は idMso でフック不能のため。詳細は下記の「重要発見」参照）
- **対象ファイル**: `.doc`（Word 97-2003）形式から開始するケース。
- **重要発見（実機検証で判明）**:
  1. **「変換」操作はファイル拡張子を変えない**：メモリ内のドキュメント形式（`CompatibilityMode`）だけ `15` にアップグレードし、ディスク上のファイル名は `.doc` のまま残る。明示的に保存し直すまで `.docx` にならない。
  2. **`UpgradeDocument` は idMso 経由で onAction フックが発火しない**：Ribbon.xml に正しく登録（リビルド・再インストール済み）してもログが記録されない。これは Office 仕様の制約。プロジェクト6 で判明した `TableSplitTable` / `TableSplitCells` と同じカテゴリ。
- **判定アプローチ（プランE: State-only Strict）の根拠**:
  - 上記1により、**`.doc` + `CompatibilityMode == 15` という組み合わせは「変換」ボタン押下を一意に示すシグネチャ**になる。
  - 他の経路では到達できない：
    - 何もしない → `.doc` + CompatMode=11
    - 「名前を付けて保存」→ .docx → `.docx` + 15（拡張子で弾かれる）
    - 「変換」+ Ctrl+S → `.docx` + 15（拡張子で弾かれる ＝ CSV手順外の余計な保存）
  - よってログAND判定は不要。状態のみで誤学習パターン（Save As .docx）を排除できる。
- **判定ポイント（実装版）**:
  ```csharp
  string ext = System.IO.Path.GetExtension(document.FullName).ToLowerInvariant();
  if (ext != ".doc") return false;
  return (int)document.CompatibilityMode == (int)WdCompatibilityMode.wdWord2013;
  ```
- **Ribbon.xml への追加（残置）**:
  - `<command idMso="UpgradeDocument" onAction="CommandOnAction" />` は **登録を残置**（コメントで「現在フック不能、将来のOffice対応に備えて」と注記）。
  - 7-1 の採点ロジックは **この登録に依存しない**。
- **既存実装からの変更点**:
  - `fileStateCheck = true` の擬似実装を **拡張子＋CompatibilityMode の実体判定** に置き換え
  - `LogReader.HasCommandExecuted("UpgradeDocument")` の AND 判定を**削除**（フック不能のため）
  - `new Application()` フォールバック起動を削除（プロジェクト6 の方針）
  - デバッグ出力を削除（クリーンな実装）
- **判定マトリクス（最終）**:
  | 操作 | 拡張子 | CompatMode | 判定 |
  | :--- | :--- | :--- | :--- |
  | 何もしない | `.doc` | 11 | × |
  | **「変換」のみ（正解操作）** | **`.doc`** | **15** | **○** ← ユニークシグネチャ |
  | 名前を付けて保存→.docx（誤学習） | `.docx` | 15 | × |
  | 「変換」+ 上書き保存（拡張子変更承諾） | `.docx` | 15 | × （CSV手順外の余計な操作） |
  | 変換途中でキャンセル | `.doc` | 11 | × |
- **検討した代替案（不採用）**:
  - **D-1**: `FileConvert` / `FileInfoConvert` など別の idMso を試す → プランEで状態判定が機能するため不要
  - **D-2**: `ThisAddIn` の `DocumentBeforeSave` / `DocumentChange` イベントで「変換時の特徴」を捕捉してログ記録 → 過剰実装
  - **D-3**: 状態のみ判定 → **これがプランEの実体**。採用。

### 7-2 文書プロパティ：会社名「ラビット出版」

- **ステータス**: **実装完了（COM厳密一致 + `.Value` 経由）**
- **判定対象（確定）**:
  - **COM: `document.BuiltInDocumentProperties["Company"]`**
- **論点と結論**:
  - 既存実装の最大の問題は **`Contains` による部分一致**：
    - 「ラビット出版社」「○○ラビット出版」など、誤った値でも合格してしまう。
    - 「ラビット出版」と完全一致したときのみ合格にすべき。
  - 値の比較は **完全一致（`==`）** とする。前後・中央空白も全角/半角の差もすべて区別する（厳密判定）。
  - プロジェクト6 の 6-3 で確立した「**緩い判定は誤学習を生む**」の思想を継承。
- **重大バグ（修正完了）**: `BuiltInDocumentProperties["Company"]` は **`DocumentProperty` COMオブジェクト** を返す。
  - 直接 `.ToString()` するとオブジェクトの**型名**（例: `System.__ComObject`）が返ってきてしまい、実値（"ラビット出版"）と比較しても**常に false**。
  - 同一プロジェクト内の `PowerPointChecker1_10.cs` では `documentProperty.Value.ToString()` のように **`.Value` 経由で取得**しており、これが**正規の作法**。
  - 既存実装の `Contains` で「偶然」合格していた経緯は要調査だが、現行のWord/.NET環境では `.Value` 経由が必須であることを実測で確認。
- **判定ポイント（実装版）**:
  ```csharp
  // ★ .Value 経由で実値を取得（型名ではなく文字列値を返す）
  dynamic companyProp = ((dynamic)document.BuiltInDocumentProperties)["Company"];
  string company = companyProp?.Value?.ToString() ?? "";
  return company == "ラビット出版";
  ```
- **既存実装からの変更点**:
  - `Contains("ラビット出版")` → **完全一致（`==`）** に変更
  - `["Company"]?.ToString()` → **`["Company"]?.Value?.ToString()`**（バグ修正）
  - `Trim()` は **入れない**（前後空白も誤入力として弾く）
  - 全角/半角の正規化（`CompareOptions.IgnoreWidth` 等）は **入れない**（厳密判定）
  - `new Application()` フォールバック起動を削除（プロジェクト6・7-1 と同方針）
- **判定マトリクス（実機検証済み）**:
  | 入力値 | 判定 | 実機結果 | コメント |
  | :--- | :--- | :--- | :--- |
  | 空欄 | × | - | 状態不一致 |
  | **`ラビット出版`** | **○** | **○ 確認済** | 正解（完全一致） |
  | `ﾗﾋﾞｯﾄ出版`（半角カタカナ） | × | **× 確認済** | 半角入力は弾く |
  | `ラビット出版　`（末尾全角空白） | × | **× 確認済** | 厳密一致 |
  | ` ラビット出版`（先頭空白） | × | - | 厳密一致 |
  | `ラビット 出版`（中央空白） | × | - | 厳密一致 |
  | `ラビット出版社` | × | - | 部分一致誤学習を弾く |
  | `○○ラビット出版` | × | - | 部分一致誤学習を弾く |
  | `らびっと出版`（ひらがな） | × | - | カナ種別は区別 |

#### 学び（重要）

- **COMの `DocumentProperty` は必ず `.Value` 経由でアクセスすること**。
  - `BuiltInDocumentProperties["Company"]` 自体は **オブジェクト**で、`.ToString()` は**型名を返す**（実測で確認）。
  - これはプロジェクト6 完了レポート（`WordChecker1_6_完了レポート.md`）にも追記済み。
- 同種のプロパティ（`Subject`, `Author`, `Title`, `Keywords` 等）を扱う今後のタスクでも、必ず `.Value` 経由で取得する。
- **不可視文字・正規化問題の検証ツールチップ**: 値が想定通り入っているのに合格しない場合は、文字数と各文字のコードポイント（U+XXXX）を一時的に出力すると、末尾の全角空白（U+3000）・BOM（U+FEFF）・改行コードなどを即座に特定できる。

### 7-3 「インテグラル」ヘッダーの挿入

- **ステータス**: **ダンプ採取フェーズ実装済み**（判定方針=「中」XMLのみ で確定、識別タグを実機ダンプで採取中）
- **方針確定事項**（論点1・2 の決定）:
  - **論点1**：全セクションを走査し、**いずれかのセクション**にインテグラル識別タグがあれば ○（Aパターン）
  - **論点2**：**「中」（XML のみ、識別タグで判定）** を採用
    - 理由：ログ AND はヘッダー種別（Integral vs Austin 等）を判別できず、識別力に貢献しない
    - 6-3 判例（同結果なら経路問わず ○）と整合
    - `HeaderInsertGallery` のフック可否未検証であり、XML識別タグで弾く方が確実
    - ステップ1（実機ダンプ）の結果、識別タグが見つからない場合のみ「厳しめ」を再検討
  - **論点3**：デバッグ方式A（`CheckTask_1_7_03` 内に一時ダンプコードを仕込む）
- **判定対象（候補）**:
  - **XML: `word/header*.xml`** 内の Building Block 識別子
  - 候補となる識別タグ：
    - `<w:docPartGallery w:val="..."/>`
    - `<w:docPartUnique/>`
    - SDTタイトル属性（`w:alias` や `w:tag` に英語名「Integral」が含まれる可能性）
    - テーマカラー・特定の図形要素（最終手段）
- **論点と結論**:
  - 「インテグラル」ヘッダーの本体は **意匠的なデザイン（罫線・図形）のみ**で、**「インテグラル」というテキスト文字列はヘッダー内に存在しない**。
  - 既存の `Contains("インテグラル")` は **原理的に常に false**。
  - **実機で正解操作を行った後の `header1.xml` をダンプして、識別可能な特徴量を採取する**必要がある。
- **改修ステップ**:
  1. ✅ デバッグコードを `CheckTask_1_7_03` に一時的に仕込む（全セクション×3種類のヘッダー XML を `C:\temp\1_7_03_header_dump.xml` に保存）
  2. 🔄 実機で **3パターン** を採取する（採取フェーズ中は常に false が返る）:
     - **A. 何もしない状態**（ベースライン） → `1_7_03_none.xml` にリネーム保存
     - **B. インテグラル挿入**（正解の特徴量） → `1_7_03_integral.xml` にリネーム保存
     - **C. 別ヘッダー（オースティン等）挿入**（誤合格させたい比較対象） → `1_7_03_other.xml` にリネーム保存
  3. ⏳ ダンプしたXMLから「B にあって C にない特徴」を確定（識別タグの優先順位は上記「候補」のとおり）
  4. ⏳ 判定ロジック実装
  5. ⏳ デバッグコード削除
- **採取フェーズの実装内容**（`C:\temp\1_7_03_header_dump.xml`）:
  - 全セクションの `Primary` / `FirstPage` / `EvenPages` の3種類のヘッダーをすべてダンプ
  - 各ヘッダーの `Range.Text`、`Range.WordOpenXML`、`Exists` プロパティをコメント付きで保存
  - デバッグログは `C:\temp\1_7_03_debug.txt` に時刻付きで出力
- **判定ロジック実装案**（ステップ4で採取結果に応じて確定）:
  ```csharp
  // 仮：採取結果が「<w:alias w:val="Integral"/>」だった場合の例
  string fullXml = document.WordOpenXML;
  // header*.xml ブロック内に Integral 識別子が含まれているか
  return Regex.IsMatch(fullXml, @"w:val\s*=\s*""Integral""", RegexOptions.IgnoreCase)
      || Regex.IsMatch(fullXml, @"<w:docPartGallery[^>]*Integral", RegexOptions.IgnoreCase);
  ```
  5. デバッグコードを削除

### 7-4 「朗読会.txt」として保存

- **ステータス**: 方針案提示済み・未実装
- **判定対象（推奨）**:
  - **ファイルシステム: 元ファイル同一フォルダ内の `朗読会.txt` の存在**
  - **拡張子 `.txt` の妥当性**（軽量検証）
  - **ログAND: `FileSaveAs` 実行ログ**（誤学習防止）
- **論点と結論**:
  - 元の `.docx` の COM/XML では別ファイルとして保存された内容を検証できない → **ファイルシステム走査が必須**。
  - 保存先は **元ファイルと同じフォルダ**を既定とする（MOS模擬アプリの作問環境では、`Wordtasks/Doc/` 配下の特定フォルダが想定される）。
  - ログAND を入れる理由：手動で `朗読会.txt` を別経路で作成しても合格してしまうリスクを排除するため。
- **判定ポイント（実装案）**:
  ```csharp
  string dir = System.IO.Path.GetDirectoryName(filePath);
  string targetPath = System.IO.Path.Combine(dir, "朗読会.txt");
  bool fileExists = System.IO.File.Exists(targetPath);
  bool hasLog = LogReader.HasCommandExecuted("FileSaveAs");
  return fileExists && hasLog;
  ```
- **既存実装からの変更点**:
  - `fileStateCheck = true` 固定を **`File.Exists` 実装**に置き換え
  - ログ判定はそのまま AND で残す

### 7-5 マクロ有効文書として保存＋読み取りパスワード「abc」

- **ステータス**: 方針案提示済み・未実装
- **判定対象（推奨）**:
  - **ファイルシステム: 元ファイル同一フォルダ内に同名 `.docm` ファイルが存在**
  - **`.docm` が暗号化されているか**（読み取りパスワード設定の証拠）
    - 暗号化された OOXML は **ZIP として開けない**（CFBF 形式になる）
    - `System.IO.Compression.ZipFile.OpenRead` で例外が出るかで判定可能
  - **ログAND: `FileSaveAs` 実行ログ**
- **論点と結論**:
  - パスワード「abc」での復号テストは **副作用（既存セッションの上書きやファイルロック）の懸念**があるため避ける。
  - 「暗号化されている＝何らかのパスワードが設定されている」という消極的判定で十分（MOSの厳密な「abc 一致」検証は外部から困難）。
  - **代替案**: `Documents.Open(path, PasswordDocument: "abc")` を read-only で試して成功するか判定する案もあるが、Word のインスタンスを汚すリスクがあり推奨しない。
- **判定ポイント（実装案）**:
  ```csharp
  string dir = System.IO.Path.GetDirectoryName(filePath);
  string baseName = System.IO.Path.GetFileNameWithoutExtension(filePath);
  string targetPath = System.IO.Path.Combine(dir, baseName + ".docm");
  if (!System.IO.File.Exists(targetPath)) return false;
  bool isEncrypted;
  try { using (var _ = System.IO.Compression.ZipFile.OpenRead(targetPath)) { isEncrypted = false; } }
  catch { isEncrypted = true; }
  bool hasLog = LogReader.HasCommandExecuted("FileSaveAs");
  return isEncrypted && hasLog;
  ```
- **既存実装からの変更点**:
  - `fileStateCheck = true` 固定を **`File.Exists` + ZIP オープン試行**に置き換え
  - ログ判定はそのまま AND で残す
- **注意**:
  - 「ファイル名」はCSV上明示されていないため、**元ファイルと同一の baseName** で `.docm` 化されることを前提とする。
  - 実機で正解操作を行った時の **既定ファイル名**を要確認。違うようなら判定条件を緩める。

---

## 技術リファレンス（XMLタグ名／API）

### 互換モード関連
- COM: `Document.CompatibilityMode`（`Microsoft.Office.Interop.Word.WdCompatibilityMode` 列挙型）
  - `wdWord2003` = 11, `wdWord2007` = 12, `wdWord2010` = 14, `wdWord2013` = 15
- XML: `word/settings.xml` 内の `<w:compat>` および `<w:compatSetting w:name="compatibilityMode" w:val="..."/>`

### 文書プロパティ関連
- COM: `document.BuiltInDocumentProperties["Company"]`（dynamic 経由でアクセス）
- XML: `docProps/app.xml` 内の `<Company>...</Company>` 要素（COMアクセスのほうがシンプル）

### ヘッダー関連
- COM: `document.Sections[i].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary]`
- XML: `word/header1.xml`（複数セクションの場合は `header2.xml`, `header3.xml`...）
- ビルトイン Building Block の識別: `<w:sdt>` 配下の `<w:docPartGallery w:val="..."/>`、`<w:docPartUnique/>`、`w:alias`／`w:tag` 属性内の英語名、テーマ色・図形要素

### 別名保存／パスワード保護
- 標準 .NET API: `System.IO.File.Exists`, `System.IO.Path.GetDirectoryName`, `System.IO.Path.Combine`
- パスワード保護判定: `System.IO.Compression.ZipFile.OpenRead`（暗号化されていれば例外）

---

## 実装上の注意点（プロジェクト7固有）

> プロジェクト6の知見（Regex Singleline／キーワードに章番号を含めない／距離制限／Ribbon.xml優先／COM安全策／デバッグ手法／効果系の注意／既存実装の改修判断）はすべて継承する。
>
> 以下はプロジェクト7で **新たに留意が必要な点**。

### ファイル状態 vs 文書状態 の区別
- プロジェクト6までは「開いている文書のXML」が判定対象だった。
- プロジェクト7では **ファイルシステム上の別ファイル**（保存先の `.txt`／`.docm`）や **元ファイル全体のメタ情報**（互換モード・文書プロパティ）が対象になる。
- 開いている `Document` オブジェクトの XML だけでは判定できないケースが多い。

### 完全一致 vs 部分一致
- プロパティ値（会社名等）の判定は **`Contains` ではなく完全一致**を基本とする。
- 「○○ラビット出版」「ラビット出版社」のような誤入力を合格にしないため。
- 前後空白の許容のために `Trim()` は推奨。

### Building Block（ビルトインギャラリー要素）の識別
- 「インテグラル」「ファセット」「オースティン」などの組み込みヘッダー／表紙／目次は、それぞれ固有の Building Block ID で参照される。
- ヘッダー XML には **テーマカラー・SDT・特定の図形要素** が含まれる。テキスト文字列だけでは識別できない場合が多いため、**XML構造の特徴量を実機で採取して判定条件を組み立てる** こと。
- プロジェクト6 の 6-4 での「正解動作後の XML を `C:\temp\xxx.xml` にダンプして比較」する手法をそのまま適用する。

### 保存系タスクの判定戦略
- 「名前を付けて保存」は `Documents.Open` のような副作用テストを伴うと不安定になりやすい。
- 基本戦略：
  1. **ファイル存在チェック**（`File.Exists`）
  2. **拡張子・サイズなど軽量な検証**
  3. **暗号化判定は ZIP オープン試行で消極的に**
  4. **必要に応じてログ AND**（`FileSaveAs` 等）
- パスワード保護判定は「ZIP として開けない＝暗号化されている」という消極的な判定で十分なケースが多い。

### VSTO ログコマンドの確認
- `MOSapp/MOS Word app/Wordtasks/md/VSTO_ログコマンド一覧.md` に既存の登録済みコマンドがまとまっている。
- 改修中に新たに必要になった `idMso` は、必ず `Ribbon.xml` への登録 → ビルド → 再インストール → ログ書き出し確認、の手順を踏む（プロジェクト6の知見）。
- 現状すべて × の原因として、**該当の `idMso` が Ribbon.xml に登録されていない可能性**を 7-4, 7-5 着手時に確認すること。

### idMso フック不能コマンド（プロジェクト6・7 で判明）
- `TableSplitTable`（プロジェクト6）
- `TableSplitCells`（プロジェクト6）
- **`UpgradeDocument`（プロジェクト7-1 で判明）**：Ribbon.xml に登録しビルド・再インストールしても `onAction` が発火せず、ログに記録されない。これは Office 仕様の制約。
- これらの場合は、**状態のユニークシグネチャ**（XML/COMの一意な組み合わせ）で代替する。7-1 では `ext=.doc + CompatibilityMode=15` が該当。

### 既存実装の `fileStateCheck = true` 固定について
- 7-1, 7-4, 7-5 の既存実装は **ファイル状態の検証が未実装の状態でログAND判定を組んでしまっている** ため、これは「改修対象（誤実装）」に該当する。実体化する。

### COM のフォールバック起動を排除
- 既存実装にある `catch { wordApp = new Application(); wordApp.Visible = true; }` のフォールバックは、**採点中に新規 Word プロセスが立ち上がるリスク**があるため削除する。
- プロジェクト6 の改修で全タスクから既に排除済み。プロジェクト7 でも同じ方針。

---

## プロジェクト7 改修の主な成果

> 各タスクの改修完了後にここへ追記。現時点では未着手のため空欄。

1. **判定精度の向上**:（未記入）
2. **安定性の向上**:（未記入）
3. **ナレッジの体系化**:（未記入）

---

## 改修ワークフロー

各タスクは以下のステップで進める：

1. **現状確認**: 既存実装（`WordChecker1_7.cs` の該当メソッド）と実機での挙動を確認。
2. **方針相談**: 本ドキュメントの「方針詳細」をベースにユーザーと方針確定。
3. **正解動作の採取**: 必要に応じて Word で正解操作を実行し、XML／ファイル状態／ログ を採取（特に 7-3）。
4. **実装**: 判定ロジックを実装。プロジェクト6の知見（Regex Singleline、COM安全策、章番号除外、完全一致厳密化）を遵守。
5. **動作確認**: 正解パターン／誤りパターン／別操作ルートを試して採点結果を確認。
6. **本ドキュメントへ追記**: 判定方針・判定ポイント・実装内容・落とし穴を記録。
7. **デバッグコードの除去**: 一時的なデバッグ出力は必ず削除。

---

> このドキュメントは作業中であり、各タスクの方針確定・実装完了に応じて随時更新する。すべてのタスクが完了したら `WordChecker1_7_完了レポート.md` にリネーム／再構成する想定。
