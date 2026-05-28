# Release 検証手順（PowerPoint 模擬アプリ）

配布前・Release 構成での動作確認手順。ソリューションは `MOSapp/MOSapp.slnx` を想定する。

---

## 1. 構成の整理（混在を防ぐ）

| 種類 | 表示名の例 | 実体 |
|------|------------|------|
| **模擬アプリ本体** | `Mos PowerPoint Mogi App` | WPF exe（`bin\Debug` / `bin\Release` / ClickOnce `publish`） |
| **VSTO 直インストール** | `PowerPointAddIn1`（.vsto の product 名） | `PowerPointAddIn1\bin\{Debug\|Release}\PowerPointAddIn1.vsto` |
| **MSI（powerpointvstosetup）** | プログラム名 `powerpointvstosetup`、COM 表示名 **`PowerPointMosVsto`** | `C:\Program Files\Rabbit\powerpointvstosetup\` 配下 |

**重要**: Debug の .vsto と Release の MSI を同時に入れると、古い DLL が残り採点がずれる（証跡が出ない・破壊検知が誤る等）。Release 試験前は **VSTO はどちらか一方だけ** に揃える。

---

## 2. クリーンアップ（PowerPoint 終了後）

1. **PowerPoint を完全終了**（タスクマネージャで `POWERPNT` が無いこと）。
2. **設定 → アプリ → インストール済みアプリ** で次があればアンインストール:
   - `powerpointvstosetup`
   - `Mos PowerPoint Mogi App`（ClickOnce で入れた場合）
3. PowerPoint → **ファイル → オプション → アドイン → COM アドイン** で、次があれば無効化または削除:
   - `PowerPointMosVsto`（MSI 由来）
   - `PowerPointAddIn1`（.vsto 直インストール由来）
4. （任意・確実）**VSTOInstaller** で直インストール分を解除:

```powershell
$vstoInstaller = "${env:CommonProgramFiles}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
$repoRoot = "（リポジトリの Mos PowerPoint Mogi App フォルダへのパス）"
@(
  "$repoRoot\PowerPointAddIn1\bin\Debug\PowerPointAddIn1.vsto",
  "$repoRoot\PowerPointAddIn1\bin\Release\PowerPointAddIn1.vsto",
  "$repoRoot\publish\PowerPointAddIn1\PowerPointAddIn1.vsto",
  "C:\Program Files\Rabbit\powerpointvstosetup\PowerPointAddIn1.vsto"
) | ForEach-Object {
  if (Test-Path $_) { & $vstoInstaller /Uninstall $_ /Silent 2>$null }
}
```

存在するパスだけ実行すればよい。

---

## 3. Visual Studio で Release ビルド

1. 構成を **Release | Any CPU** にする。
2. **ソリューションのビルド**（`MOSapp.slnx`）。
   - `PowerPointAddIn1` → `bin\Release\`
   - `MOS PowerPoint app` → `bin\Release\Mos PowerPoint Mogi App.exe`
   - `powerpointvstosetup` → `MOSapp\powerpointvstosetup\Release\setup.exe`（vdproj は Release の AddIn 出力を参照）

**確認ポイント**

- `PowerPointAddIn1\bin\Release\PowerPointAddIn1.vsto` の更新日時がビルド直後であること。
- `powerpointvstosetup.vdproj` が `bin\Release\PowerPointAddIn1.vsto` を取り込む設定であること。

---

## 4. VSTO を配布と同じ方法でインストール（推奨）

配布想定どおり試す場合:

```
MOSapp\powerpointvstosetup\Release\setup.exe
```

（または同フォルダの `powerpointvstosetup.msi`）

- インストール先の目安: `C:\Program Files\Rabbit\powerpointvstosetup\`
- PowerPoint の COM アドイン一覧では **`PowerPointMosVsto`** として登録される。

インストール後、**PowerPoint を起動**しアドインが有効か確認する。

---

## 5. 模擬アプリ本体を Release で起動

Release 試験中は **Debug の exe を使わない**。

```
Mos PowerPoint Mogi App\bin\Release\Mos PowerPoint Mogi App.exe
```

Visual Studio から起動する場合も、スタートアッププロジェクトの構成が **Release** であること。

---

## 6. 動作確認（最低限）

| 確認項目 | 内容 |
|----------|------|
| ログ依存タスク | **5-1** / **10-4** / **11-7** をその場採点・一括採点の両方で ○ |
| 証跡ログ | `%TEMP%\mos_ppt_task_evidence.txt` に `[Task5-1] Print` / `[Task10-4] Grayscale` / `[Task11-7] Print` |
| アドイン | COM アドインで **PowerPointMosVsto** が有効（MSI 経路の場合） |

その他タスク（1-3 / 1-4 / 3-4 / 9-6 等）の再発時は [配布版_不合格タスク精査メモ_20260528.md](./配布版_不合格タスク精査メモ_20260528.md) を参照。

---

## 7. ClickOnce「発行（publish）」で配布全体を試す場合（任意）

本番に近い流れ:

1. `MOS PowerPoint app` を **発行（Publish）** → `Mos PowerPoint Mogi App\publish\setup.exe`
2. その **setup.exe** で模擬アプリ本体をインストール。
3. VSTO は **次のどちらか一方だけ**（両方入れない）:
   - **A（配布 MSI・推奨）**: 手順 4 の `powerpointvstosetup\Release\setup.exe`
   - **B（発行フォルダの .vsto）**: `publish\PowerPointAddIn1\PowerPointAddIn1.vsto` をダブルクリック

`AfterPublish` により AddIn ファイルは `publish\PowerPointAddIn1\` にコピーされるが、**MSI と .vsto 直インストールは別登録**のため併用しない。

---

## 8. 開発時の使い分け

| 目的 | 模擬アプリ | VSTO |
|------|------------|------|
| 日常開発 | Debug exe | `bin\Debug\PowerPointAddIn1.vsto`（`scripts\Deploy-PowerPointVsto.ps1` 等） |
| Release / 配布確認 | Release exe または publish の setup | **`powerpointvstosetup` の setup.exe** |
| 切り替え時 | PowerPoint 終了 → 古い VSTO をアンインストール → 新方式をインストール → PowerPoint 再起動 | 同左 |

---

## 関連ドキュメント

- [配布版_不合格タスク精査メモ_20260528.md](./配布版_不合格タスク精査メモ_20260528.md) — 不合格報告タスクの原因候補と対応状況
- [PPアドイン_ログタブ表示の解消要因.md](./PPアドイン_ログタブ表示の解消要因.md) — .vsto の再インストール手順
- [PP採点_VSTOで検証する問題一覧.md](./PP採点_VSTOで検証する問題一覧.md) — ログ依存タスク一覧
