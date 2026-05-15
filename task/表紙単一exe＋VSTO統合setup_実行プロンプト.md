# 表紙単一exe＋VSTO統合 setup — 実行プロンプト

> **使い方**: このドキュメント全体をプロンプトにコピーして送信すると、記載の手順をそのまま実行できる。

---

## 実行依頼（プロンプト冒頭用）

以下を実施してください。

1. **統合ブートストラップ**の C# プロジェクトを新規作成し、発行後に `publish\setup.exe` として配置する。
2. **発行スクリプト**（`Publish-MOSapp-With-VSTO.ps1`）を修正し、表紙の setup をリネームしたうえでブートストラップを `publish\setup.exe` に配置する。
3. フォルダ名の定数化（`WordAddIn` / `ExcelAddIn` / `PowerPointAddIn`）をブートストラップと発行スクリプトで揃える。

前提・現状・詳細仕様は以下に従ってください。

---

## 前提

- 既存プラン「表紙のみ表示・単一exe構成」は実施済みとする（ルートは表紙のみ、3科目は `App` サブフォルダ、Merge 先は `App`、表紙の exe パス解決はサブフォルダ優先）。
- **1 つの setup.exe を実行するだけで、表紙アプリと VSTO（Word／将来 Excel・PowerPoint）がすべてインストールされる**ようにする。
- VSTO は Word 既存、Excel 製作中、PowerPoint 計画あり。

---

## 現状の整理

| 対象 | 現状 |
|------|------|
| 表紙 | `MOSapp/MOSapp/MOSapp.csproj` で ClickOnce 発行 → `publish\` に setup.exe と Application Files |
| 3科目 | `Merge-AppOutputs-ToPublish.ps1` で `publish\App\` と Application Files 内にマージ済み |
| Word VSTO | `Publish-MOSapp-With-VSTO.ps1` でビルドし `publish\WordAddIn\` にコピー。ユーザーは別途 .vsto 実行が必要 |
| Excel/PPT VSTO | 未作成または製作中。同じ発行フローに載せる想定 |

---

## 方針: 統合ブートストラップ（単一 setup.exe）

ユーザーが実行する **setup.exe = 統合ブートストラップ** とし、1 本で次を順に実行する。

1. **表紙の ClickOnce インストール**（表紙用 setup を起動し、完了を待つ）
2. **Word VSTO のインストール**（`VSTOInstaller.exe /i "…\WordAddIn\*.vsto" /s`）
3. **将来** Excel VSTO・PowerPoint VSTO も同様（フォルダ・プロジェクトができたらステップ追加）

- 配布フォルダルートの「メイン setup.exe」をブートストラップに差し替える。表紙が生成する ClickOnce の setup は **`Setup_MOSapp_ClickOnce.exe`** にリネームして保存し、ブートストラップから呼び出す。
- .vsto パスは **ブートストラップと同じ配布フォルダ** を基準（例: `(setup.exe のフォルダ)\WordAddIn\New_MOSWordVSTOAddIn.vsto`）。`publish\WordAddIn\` にコピーしている構成はそのまま利用。

---

## 実施内容（変更対象）

### 1. 統合ブートストラップの作成

- **新規**: 統合 setup 用の小さい C# プロジェクト（コンソール or WPF）をリポジトリに 1 つ追加。発行後に `publish\` に配置する exe を出力する。
- **処理**:
  1. 自分（exe）のディレクトリを `baseDir` とする。
  2. 表紙の ClickOnce を実行: `Process.Start(baseDir + "Setup_MOSapp_ClickOnce.exe")` で起動し、`WaitForExit` で完了を待つ。
  3. `VSTOInstaller.exe` を探す（`%CommonProgramFiles%\microsoft shared\VSTO\10.0\VSTOInstaller.exe` 等）。見つかれば `VSTOInstaller.exe /i "baseDir\WordAddIn\New_MOSWordVSTOAddIn.vsto" /s` を実行（終了待ち）。
  4. （将来）`ExcelAddIn\*.vsto`、`PowerPointAddIn\*.vsto` が存在すれば同様に `/i` と `/s` で実行。
- **証明書**: サイレント（`/s`）は Trusted Publishers に証明書が必要。ブートストラップの「最初のステップ」で `certutil -addstore TRUSTEDPUBLISHER ...` を呼ぶオプションを用意するか、README で「初回のみ .vsto を手動実行して信頼」を案内する。

**定数**: フォルダ名は `WordAddIn` / `ExcelAddIn` / `PowerPointAddIn` を 1 か所で定義し、将来の追加時に揃える。

### 2. 発行フロー側の変更

- **`Publish-MOSapp-With-VSTO.ps1`**（または `Publish-MOSapp-Full.ps1` から呼ばれる流れ）:
  - 表紙を Publish したあと、生成された `publish\setup.exe` を **`Setup_MOSapp_ClickOnce.exe`** にリネームする。
  - 統合ブートストラップの exe をビルドし、**`publish\setup.exe`** として配置する（上書き）。
- **`Merge-AppOutputs-ToPublish.ps1`**: 変更不要（既に `App` サブフォルダへマージ済み）。ブートストラップは配布フォルダの `WordAddIn` を参照するので、Application Files 内への WordAddIn コピーは必須ではない。

### 3. ClickOnce と VSTO

- 今回は「配布フォルダをそのまま渡す」前提。ブートストラップは **配布フォルダ内の WordAddIn** を参照する形で十分。配布フォルダを消さないよう README で案内する。

### 4. 定数・名前の統一

- サブフォルダ名: 表紙の `SubjectAppsSubfolderName = "App"` と `Merge-AppOutputs-ToPublish.ps1` の `$SubjectAppsSubfolderName = "App"` は既に一致。変更なし。
- VSTO 用フォルダ名: ブートストラップと発行スクリプトで `WordAddIn` / `ExcelAddIn` / `PowerPointAddIn` を定数または 1 か所定義し、将来の追加時に揃える。

---

## 動作確認のポイント

- **開発時**: 表紙から科目選択し、`App` 内の 3 科目 exe が起動すること（既存どおり）。
- **発行後**:
  - ルートに **setup.exe が 1 つ**（統合ブートストラップ）と、`Setup_MOSapp_ClickOnce.exe`（表紙用）、`App\`、`WordAddIn\` などがあること。
  - **setup.exe を 1 回実行**すると、表紙の ClickOnce インストールのあと、Word VSTO が VSTOInstaller /s でインストールされること。
- **証明書**: 他 PC でサイレント VSTO が失敗する場合は、Trusted Publishers に証明書を入れるか、README で .vsto を 1 回手動実行する手順を用意する。

---

## 補足（VSTO サイレントインストール）

- コマンド: `VSTOInstaller.exe /i "（.vsto の絶対パス）" /s`
- VSTOInstaller の典型パス: `%CommonProgramFiles%\microsoft shared\VSTO\10.0\VSTOInstaller.exe`
- `/s` はサイレント。証明書が Trusted Publishers にないと失敗するため、配布時は証明書の扱いを README またはブートストラップのオプションで明示する。

---

## 実行チェックリスト（実装時）

- [ ] 統合ブートストラップ用 C# プロジェクトを新規作成（リポジトリ内の適切な場所）
- [ ] ブートストラップで baseDir 取得 → Setup_MOSapp_ClickOnce.exe 実行・待機
- [ ] VSTOInstaller 検索・WordAddIn\New_MOSWordVSTOAddIn.vsto を /i … /s で実行
- [ ] フォルダ名定数（WordAddIn / ExcelAddIn / PowerPointAddIn）を定義
- [ ] Publish-MOSapp-With-VSTO.ps1: 表紙 Publish 後、publish\setup.exe → Setup_MOSapp_ClickOnce.exe にリネーム
- [ ] Publish-MOSapp-With-VSTO.ps1: ブートストラップをビルドし publish\setup.exe に配置
- [ ] 発行後の setup.exe 1 回実行で表紙＋Word VSTO がインストールされることを確認
