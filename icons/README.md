# PNG → ICO → アプリ登録手順

透過付き PNG から Windows 用アイコンを作り、WPF アプリ（表紙／Word／Excel／PowerPoint）の EXE と ClickOnce 発行に載せる手順。

---

## 1. 全体フロー

```text
icons/pictures/*.png   （素材・透過あり）
        │
        ▼
icons/Convert-Icons.ps1
        │
        ├─► icons/<Name>.ico
        └─► <AppProject>/AppIcon.ico   （自動コピー）
                │
                ├─► .csproj ApplicationIcon
                └─► ビルド／VS 発行で EXE・スタートメニューに反映
```

表示名（例: 前田式Excel演習システム）は各 `.csproj` の `ProductName` と `AssemblyInfo` で設定する。

---

## 2. 素材 PNG の要件

| 項目 | 推奨 |
| --- | --- |
| サイズ | 512×512（正方形） |
| 形式 | PNG、`Format32bppArgb`（アルファ付き） |
| 背景 | 四隅は**完全透過 (A=0)**。不透明の黒四角にしない |
| 半透明 | 縁のアンチエイリアス（A=1〜254）があってよい |
| 配置 | `icons/pictures/` |

注意:

- 見た目が「黒い背景」でも、実データは透過のことが多い。変換前に四隅のアルファを確認する
- **黒ピクセルを一律透過にしない**（キャラの目・線画が消える）
- ファイル名に空白や日本語があってもよい（スクリプト側はワイルドカードで拾う）

---

## 3. ICO 生成（`icons/Convert-Icons.ps1`）

### 実行

リポジトリルートで:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\icons\Convert-Icons.ps1
```

プレビュー（チェッカー背景）が不要な場合:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\icons\Convert-Icons.ps1 -SkipPreview
```

### 生成されるサイズ

| サイズ | 格納形式 |
| --- | --- |
| 16 / 32 / 48 / 64 | クラシック 32bpp BMP（XOR + AND、アルファ付き） |
| 256 | PNG 圧縮フレーム |

全サイズを PNG-in-ICO だけにすると、Explorer でアイコンが欠けたりデフォルトに戻ることがある。**ハイブリッド形式**にする。

### スクリプトが行うこと

1. `icons/pictures/` から製品別 PNG をパターンマッチで取得
2. HighQualityBicubic で各サイズにリサイズ（アルファ保持）
3. `icons/Word.ico` 等を出力
4. 四隅透過・半透明ピクセルを数値検証（失敗なら終了）
5. 任意で `icons/_preview/<Name>_checker.png` を出力
6. 各 WPF プロジェクト直下へ `AppIcon.ico` をコピー（表紙は Word うさぎを共用）

### 新しいアプリ用にジョブを足す例

`Convert-Icons.ps1` の `$jobs` に追記:

```powershell
$jobs = @(
    @{ Name = 'Word';       Pattern = '*Word_0.png';       Out = 'Word.ico' }
    @{ Name = 'Excel';      Pattern = '*Excel*_0.png';     Out = 'Excel.ico' }
    @{ Name = 'PowerPoint'; Pattern = '*PowerPoint_0.png'; Out = 'PowerPoint.ico' }
    # 追加例
    @{ Name = 'MyApp';      Pattern = '*MyApp*_0.png';      Out = 'MyApp.ico' }
)
```

末尾のプロジェクト同期も追加:

```powershell
@{ Src = 'MyApp.ico'; Dest = Join-Path $root '..\MOSapp\MyAppFolder\AppIcon.ico' }
```

---

## 4. WPF アプリ本体への登録

### 必須ルール

**`ApplicationIcon` に `..\icons\....ico` のような親ディレクトリ参照を書かない。**

ClickOnce / 生成マニフェストに `..\..\icons\Word.ico` が入り、起動時に SxS エラー **14001**（サイドバイサイド構成が正しくない）になる。

### 手順

1. プロジェクト直下に `AppIcon.ico` を置く（`Convert-Icons.ps1` が自動コピー）
2. `.csproj` の先頭付近の `PropertyGroup` に追加:

```xml
<ApplicationIcon>AppIcon.ico</ApplicationIcon>
<ProductName>前田式Excel演習システム</ProductName>
<PublisherName>Rabbit</PublisherName>
```

本リポジトリの例:

| 製品 | プロジェクト | ApplicationIcon | 表示名例 |
| --- | --- | --- | --- |
| 表紙 | `MOSapp/MOSapp/MOSapp.csproj` | `AppIcon.ico`（Word うさぎ共用） | 前田式MOS演習システム |
| Word | `MOSapp/MOS Word app/MOS Word app.csproj` | `AppIcon.ico` | 前田式Word演習システム |
| Excel | `MOSapp/mos_xaml_app/MOSExcelMogiApp.csproj` | `AppIcon.ico` | 前田式Excel演習システム |
| PowerPoint | `MOSapp/Mos PowerPoint Mogi App/MOS PowerPoint app.csproj` | `AppIcon.ico` | 前田式PowerPoint演習システム |

これで EXE にアイコンが埋め込まれ、VS 発行（ClickOnce）後のスタートメニューにも反映される。

---

## 5. ビルド・確認

```powershell
# アイコン再生成（素材を差し替えたとき）
powershell -NoProfile -ExecutionPolicy Bypass -File .\icons\Convert-Icons.ps1
```

その後、対象プロジェクトをビルドまたは VS で発行する。

確認項目:

1. `icons/_preview/*_checker.png` で四隅が透過（チェッカーが見える）
2. `bin\Release\*.exe` のアイコンが製品色になっている
3. 発行してインストール後、スタートメニューの名前・アイコン
4. アプリが **SxS 14001 で落ちない**（親パスの `ApplicationIcon` をしていないこと）

古いアイコンのまま見える場合は Windows のアイコンキャッシュのことが多い。エクスプローラー再起動、または別フォルダに Setup をコピーして確認する。

---

## 6. 新しいアプリへ横展開するチェックリスト

1. [ ] `icons/pictures/` に 512×512 透過 PNG を置く
2. [ ] `Convert-Icons.ps1` の `$jobs` と `AppIcon.ico` 同期先を追加
3. [ ] スクリプト実行 → `icons/<Name>.ico` と `<Project>/AppIcon.ico` ができる
4. [ ] `.csproj` に `<ApplicationIcon>AppIcon.ico</ApplicationIcon>`（**相対で `..` を使わない**）
5. [ ] `.csproj` に `ProductName` / `PublisherName`、`AssemblyInfo` の表示名を設定
6. [ ] ビルドまたは VS 発行し、EXE・スタートメニューを目視確認

---

## 7. 今回つまずいた点（再発防止）

| 症状 | 原因 | 対策 |
| --- | --- | --- |
| 起動時に 14001（SxS） | `ApplicationIcon` が `..\..\icons\....ico` | プロジェクト直下の `AppIcon.ico` のみ参照 |
| アイコンが時々デフォルト | 全サイズ PNG-in-ICO、64px 欠落 | 16–64=BMP32、256=PNG。64 を含める |
| 透過が潰れる／黒四角 | 黒をキーアウト、またはアルファ非保持リサイズ | 既存アルファを保持。チェッカーで検証 |
| 入れ直してもアイコンが古い | Explorer キャッシュ | 再起動または別パスで確認 |
| スタートメニュー名が EXE 名のまま | `ProductName` 未設定 | ClickOnce の `ProductName` と AssemblyInfo を設定 |

---

## 8. 関連ファイル

| パス | 内容 |
| --- | --- |
| `icons/pictures/` | 素材 PNG |
| `icons/Convert-Icons.ps1` | PNG→ICO＋検証＋AppIcon 同期 |
| `icons/*.ico` | 共有 ICO |
| `icons/_preview/` | 透過確認用（gitignore 対象） |
| `MOSapp/**/AppIcon.ico` | 各アプリの ApplicationIcon |
