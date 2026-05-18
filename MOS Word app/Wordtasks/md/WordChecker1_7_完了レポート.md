# プロジェクト7 採点機能 改修完了レポート

プロジェクト7（互換モード解除／文書プロパティ／ヘッダー挿入／別名保存）全5タスクの採点ロジック改修が完了した。
本レポートは最終的な実装方針・判定ロジック・実装上の知見をまとめたものであり、プロジェクト8以降の改修における設計指針としても活用可能。

> 設計指針・判定方針の決定フロー・実装上の注意点は、[`WordChecker1_6_完了レポート.md`](./WordChecker1_6_完了レポート.md) で体系化された知見を継承する。本レポートでは **プロジェクト7固有の方針・判定ロジック・落とし穴** を記載する。

---

## 最終サマリ

| タスク | 判定方針 | 採点結果 |
| :--- | :--- | :--- |
| 7-1 | `.doc`+Compat15 **または** `UpgradeDocument` ログ（5-1 同型） | **○ 実機確認済** |
| 7-2 | `BuiltInDocumentProperties["Company"].Value` の完全一致 | **○ 実機確認済** |
| 7-3 | 全セクション×ヘッダー種別の `WordOpenXML` **または** `IntegralHeader` ログ | **○ 実機確認済** |
| 7-4 | `朗読会.txt` 存在 + **`FileSaveAsTxt`** ログ | **○ 実機確認済** |
| 7-5 | `朗読会.docm` + CFB 暗号化 + **`FileSaveAsDocm`** ログ | **○ 実機確認済** |

改修順序は **7-1 → 7-2 → 7-3 → 7-4 → 7-5** で実施した。

**実装ファイル**

- 採点: `MOSapp/MOS Word app/Libraries/WordChecker/Group1/WordChecker1_7.cs`
- VSTO: `MOSapp/MOS Word app/New_MOSWordVSTOAddIn/New_MOSWordVSTOAddIn/ThisAddIn.cs`
- ログ一覧: [`VSTO_ログコマンド一覧.md`](./VSTO_ログコマンド一覧.md)

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

## 改修前の問題点（すべて解消済み）

| タスク | 改修前 | 改修後 |
| :--- | :--- | :--- |
| 7-1 | `fileStateCheck = true` 固定 | `.doc`+15 **または** ポーリング `UpgradeDocument` |
| 7-2 | `Contains` + `.ToString()` 誤用 | `.Value` 経由の **完全一致** |
| 7-3 | ヘッダー本文に「インテグラル」文字列検索（常に false） | `WordOpenXML` 構造指紋 + `IntegralHeader` ログ OR |
| 7-4 | `fileStateCheck = true` 固定 | `朗読会.txt` + **`FileSaveAsTxt`** |
| 7-5 | `fileStateCheck = true` 固定 | `朗読会.docm` + CFB 暗号化 + **`FileSaveAsDocm`** |

---

## タスク別 最終実装サマリ

| タスク | 内容 | 判定方針 | 判定ポイント |
| :--- | :--- | :--- | :--- |
| **7-1** | 互換モードの解除 | **状態 OR ログ** | `.doc` + `CompatibilityMode == 15` **または** `UpgradeDocument`（ThisAddIn ポーリング） |
| **7-2** | 文書プロパティ：会社名 | **COM 厳密一致** | `Company.Value == "ラビット出版"`（Trim・正規化なし） |
| **7-3** | 「インテグラル」ヘッダー | **XML OR ログ** | 全セクション×Primary/First/Even の指紋 **または** `IntegralHeader` |
| **7-4** | 「朗読会」txt 保存 | **ファイル + ログ AND** | `朗読会.txt` + **`FileSaveAsTxt`** |
| **7-5** | 「朗読会」docm + 読取PW | **ファイル + 暗号化 + ログ AND** | `朗読会.docm` + CFB 先頭 + **`FileSaveAsDocm`** |

---

## 各タスクの方針詳細

### 7-1 互換モードの解除

- **判定**: `stateOk || logOk`（5-1 と同型）
- **重要**: 「変換」は拡張子を `.doc` のままにし、`CompatibilityMode` のみ 15 にする。`.doc`+15 は正解操作のユニークシグネチャ。
- **`UpgradeDocument`**: Ribbon の idMso は発火しない。**ThisAddIn ポーリング**で `.doc` 上の非15→15 遷移時に記録。7-4 後など拡張子が変わってもログがあれば再採点で ○。

```csharp
bool stateOk = ext == ".doc"
    && (int)document.CompatibilityMode == (int)WdCompatibilityMode.wdWord2013;
bool logOk = LogReader.HasCommandExecuted("UpgradeDocument");
return stateOk || logOk;
```

### 7-2 文書プロパティ：会社名「ラビット出版」

- **`DocumentProperty` は必ず `.Value` 経由**（`.ToString()` は型名が返る）
- **完全一致**（`==`）。`Trim()` や全角半角正規化は入れない。

### 7-3 「インテグラル」ヘッダーの挿入

- **`Header.Range.WordOpenXML` のみ**（Packaging 参照はビルド都合で不採用）
- 全セクション×Primary/First/Even を走査。Integral メタ、厳密指紋、`.doc`+互換15 時の補助指紋（Accent2+2列 tbl）を OR
- **7-4 後**: `TryIntegralFromLiveHeaderWordOpenXml || IntegralHeader` ログ（7-1 と同型の再採点補完）
- VSTO の `EvaluateIntegralHeaderPresenceForPolling` とチェッカー内ロジックは **二重実装**（変更時は両方を同期）

### 7-4 「朗読会.txt」として保存

- **`朗読会.txt` の存在**（教材フォルダに事前作成されていても可）
- **`FileSaveAsTxt`**: VSTO ポーリング／`DocumentBeforeSave` で **ActiveDocument が `朗読会.txt` に遷移**したときのみ（初回オープンだけではログなし）
- Ribbon の汎用 `FileSaveAs` は採点に未使用

**実機確認済み（2026-05）**: 正解操作で `FileSaveAsTxt` が記録され採点 ○。

### 7-5 マクロ有効文書＋読み取りパスワード「abc」

- **`朗読会.docm`**（7-4 の txt と同ベース名・同フォルダ）
- **暗号化**: 先頭 4 バイトが CFB `D0 CF 11 E0`（読取PW相当）。パスワードなし docm（ZIP `PK`）は ×。文字列 `abc` の一致は検証しない。
- **`FileSaveAsDocm`**: `朗読会.docm` へ遷移したときのみ。**`FileSaveAsTxt` では 7-5 合格にならない**（事前 docm が残っていても、7-4 のみでは ×）。

**実機確認済み（2026-05）**:

| パターン | 7-5 |
| :--- | :--- |
| 7-4 のみ（事前 `朗読会.docm` あり含む） | **×** |
| 7-5 正解操作（PW 付き docm 保存） | **○** |

---

## 共通の設計指針（プロジェクト6から継承）

1. **ターゲット特定**: プロジェクト7は表操作が無く、本指針の出番は限定的。
2. **VSTO 制約**: `idMso` が発火しない操作はポーリング等で補完（7-1, 7-4, 7-5）。
3. **COM 最小化**: 可能なら `WordOpenXML` / `BuiltInDocumentProperties` で完結。
4. **厳密判定**: `Contains` による緩い合格は避ける（7-2）。

### プロジェクト7 特有の判定軸

| 操作 | 方針 |
| :--- | :--- |
| 互換モード解除 | **状態 OR ログ** |
| 文書プロパティ | **COM 完全一致** |
| ヘッダー挿入 | **XML OR ログ**（7-4 後の再採点用） |
| 別名保存 | **ファイル存在 + タスク専用ログ**（`FileSaveAsTxt` / `FileSaveAsDocm` を分離） |
| パスワード保護 | **CFB 先頭マジック**（読取PW相当。`abc` 一致は見ない） |

---

## 技術リファレンス

### 互換モード
- COM: `Document.CompatibilityMode`（`wdWord2013` = 15）
- 7-1 シグネチャ: 拡張子 `.doc` + CompatMode 15

### 文書プロパティ
- `((dynamic)document.BuiltInDocumentProperties)["Company"].Value`

### ヘッダー（7-3）
- `document.Sections[i].Headers[WdHeaderFooterIndex].Range.WordOpenXML`
- 指紋例: `fill="E97132"`, `w:w="1782"`, `w:w="7286"` 等

### 別名保存（7-4 / 7-5）
- `File.Exists`, `Path.Combine(dir, "朗読会.txt"|"朗読会.docm")`
- 暗号化: 先頭 `D0 CF 11 E0`（CFB）= 読取PW付き相当。`50 4B`（PK）= パスワードなし OOXML

### VSTO ログ ID（採点で使用）

| ログ ID | タスク |
| :--- | :--- |
| `UpgradeDocument` | 7-1 |
| `IntegralHeader` | 7-3（補完） |
| `FileSaveAsTxt` | 7-4 |
| `FileSaveAsDocm` | 7-5 |

---

## 実装上の注意点（プロジェクト7固有）

### ファイル状態 vs 文書状態
- 7-4/7-5 は **ディスク上の別ファイル**を `File.Exists` で検証する。
- 採点時の `filePath` は `ActiveDocument.FullName` 由来のため、7-4 後にアクティブが `朗読会.txt` でも **同一 `dir`** で `朗読会.docm` を探せる。

### 事前作成ファイルがある運用
- 教材フォルダに **あらかじめ `朗読会.txt` / `朗読会.docm` がある**前提でもよい。
- 合格には **今回セッションで該当ファイル名へ遷移した専用ログ**が必要（事前ファイルのみでは不可）。

### 7-2 の完全一致
- プロパティは **`Contains` ではなく `==`**。7-2 では **`Trim()` も入れない**（誤入力の空白を弾く）。

### idMso フック不能（判明済み）
- **`UpgradeDocument`**: Ribbon では発火しない → ポーリングで記録。
- **Backstage の名前を付けて保存**: 汎用 `FileSaveAs` は環境により未発火 → **`FileSaveAsTxt` / `FileSaveAsDocm`** をポーリングで付与。

### COM フォールバック排除
- `new Application()` による採点時の Word 新規起動は **全タスクで削除**。

### VSTO とチェッカーの二重実装
- 7-3（インテグラル指紋）と 7-4/7-5（保存遷移ログ）は `ThisAddIn.cs` と `WordChecker1_7.cs` の両方にロジックがある。変更時は **両方を同期**すること。

---

## プロジェクト7 改修の主な成果

1. **判定精度の向上**
   - 互換モード・会社名・インテグラルヘッダー・別名保存を、誤学習しにくい状態／構造／専用ログで判定。
   - 7-4/7-5 でログ ID を分離し、「事前 docm + 7-4 のみ」で 7-5 が誤合格する問題を解消。

2. **安定性の向上**
   - `DocumentProperty.Value` バグ修正、7-4 後の 7-1/7-3 再採点（ログ OR）、Backstage 保存のポーリング補完。

3. **ナレッジの体系化**
   - `.doc`+Compat15 のユニーク性、CFB/ZIP による PW 判定、VSTO 遷移ログのベースライン設計を本レポートに集約。

---

## 既知の限界（仕様として許容）

- 7-5: パスワードが **`abc` かどうか**は検証しない（暗号化の有無のみ）。
- 7-5: **書き込みパスワードのみ**の docm は ZIP のままのことがあり × になる（課題は読取PW）。
- 7-4: ファイル内容が本当にプレーンテキストかは未検証（拡張子 `.txt` のみ）。
- 7-3: 他ビルトインヘッダーとの指紋衝突時は XML ダンプで条件追加が必要。

---

> 旧作業ドキュメント: `WordChecker1_7_改修中.md`（本レポートに統合済み）
