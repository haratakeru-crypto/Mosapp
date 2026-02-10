---
name: MOS PPT配布・Word版・CSV保存
overview: (1) 現行PowerPointアプリの他PC配布のため、データパス設定のポータブル化と配置手順の整理、(2) PowerPointアプリと同構成のWord版を新規作成、(3) 結果画面表示時に保存先を指定してCSVを自動出力する機能を追加する。
todos: []
isProject: false
---

# MOS PowerPoint アプリ：他PC配布・Word版・CSV保存 実装計画

## 1. 現状の整理

- **PowerPointアプリ** ([MainViewModel.cs](c:\Users\kouza\source\repos\MOS PowerPoint app\MainViewModel.cs), [UiTestAppBarWindow.xaml.cs](c:\Users\kouza\source\repos\MOS PowerPoint app\Views\UiTestAppBarWindow.xaml.cs))
  - データパス: `Assets\config.json` の `powerPointDataPath` → 未設定時は `App.config` → 未設定時は `C:\MOSTest\PowerPoint365`
  - プロジェクト: `Tab1`～`Tab3` 配下の `Project1.pptx`～`Project11.pptx`
  - PowerPoint Interop + Office.Core 参照。Office は GAC の `office.dll`（HintPath がマシン固有）
- **結果画面** ([ResultWindow.xaml](c:\Users\kouza\source\repos\MOS PowerPoint app\Views\ResultWindow.xaml), [ResultWindow.xaml.cs](c:\Users\kouza\source\repos\MOS PowerPoint app\Views\ResultWindow.xaml.cs))
  - 正答率・✖数・プロジェクト/タスク一覧（`ResultProjectInfo` / `ResultTaskInfo`）を表示。CSV 出力は未実装。
- **配置**: ClickOnce 形式（`bin\...\app.publish` に setup.exe、Application Files）。JSON・config は出力に同梱。

---

## 2. 他PCで使えるようにする（PowerPointアプリ）

**目的**: 配布先PCでも、データフォルダの位置を柔軟に変えて実行できるようにする。


| 項目                    | 内容                                                                                                                                                                 |
| --------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **データパス**             | (a) 起動時、`config.json` に `powerPointDataPath` が無い or そのフォルダが存在しない場合、フォルダ選択ダイアログで「PowerPoint データルート」（`Tab1`～`Tab3` の親）を選ばせる。(b) 選んだパスを `config.json` に保存し、次回からそれを読む。 |
| **config.json の置き場所** | 実行ファイルと同じフォルダの `Assets\config.json` を優先（現状どおり）。配置時も同梱する。                                                                                                           |
| **Office 参照**         | GAC の `office.dll` 依存は他PCでバージョン差異の原因になりうる。NuGet の `Microsoft.Office.Interop.PowerPoint` 等に切り替え、HintPath をやめる（検討）。難しければ、運用で「同程度の Office バージョン」を案内。                  |
| **配布物**               | ClickOnce の setup.exe + Application Files。README に「.NET 4.8」「PowerPoint インストール」「初回のデータフォルダ選択」を記載。                                                                  |


**主な変更箇所**:

- `MainViewModel.LoadProjects()`: 上記 (a)(b) の「パスが無い → フォルダ選択 → config 保存」ロジックを追加。
- （Optional）`MOS PowerPoint app.csproj`: Office を NuGet 参照に変更。
- ルートに `README.md` または `配布手順.txt` を追加し、他PCでのセットアップ手順を明記。

---

## 3. Word版の新規作成（PowerPoint版と同構成）

**目的**: 現在の PowerPoint アプリと同様の機能を Word で実現する。

- **新規ソリューションプロジェクト**: 例 `MOS Word app`（.NET 4.8 WPF）。ソリューションディレクトリは `MOS PowerPoint app` と並列などを想定。
- **構成の踏襲**:
  - `MainWindow` / `MainViewModel`: プロジェクト一覧・「開く」で Word 起動。
  - `Views\UiTestAppBarWindow`: アプリバー。Word ウィンドウ配置（`PositionWordWindow`）、タスク表示、閉じる/レビュー → 結果画面。
  - `Views\ResultWindow`: 結果表示（後述 CSV 連携も同じ方針で実装）。
- **Word 固有の差し替え**:
  - Interop: `Microsoft.Office.Interop.Word`。`Document` を開く・取得。プロセス名 `WINWORD`、ウィンドウクラス `OpusApp`（[Word_PowerPoint版作成ガイド](c:\Users\kouza\source\repos\MOS PowerPoint app\mos_xaml_app\tasks\Word_PowerPoint版作成ガイド.md) 準拠）。
  - データパス: `wordDataPath`（`config.json`）→ `C:\MOSTest\Word365` 等。`Tab1`～`Tab3`、`Project1.docx`～`Project11.docx`。
  - 問題文 JSON: `MOS模擬アプリ問題文一覧_Word.json` を用意し、`LoadTasks` / `ResultWindow` で参照。
- **リセット・次のプロジェクト等**: PowerPoint 版の `UiTestAppBarWindow` の処理を Word 用に置き換え（`CloseAllPowerPointPresentations` → `CloseAllWordDocuments`、パス・拡張子の切り替え）。

**成果物**:

- 新規 csproj、App / MainWindow / MainViewModel、`UiTestAppBarWindow`、`ResultWindow`、`Assets\config.json`、`References\JSON\MOS模擬アプリ問題文一覧_Word.json`（中身は暫定でOK）。

---

## 4. 結果画面でCSV保存（PowerPoint版・Word版共通）

**要件**: 結果表示後に、保存先をユーザーが指定し、**自動で1回** CSV 出力する。

**フロー**:

1. 試験終了 → 結果画面を表示（既存の `ShowResultWindow` 等の流れはそのまま）。
2. 結果画面の `Loaded` でデータ読み込み・表示が終わったあと、**保存先ダイアログ**（`SaveFileDialog`）を出す。初期ファイル名例: `MOS結果_YYYYMMDD_HHmmss.csv`。
3. ユーザーが保存先を指定して OK → そのパスに CSV を書き出し。キャンセルした場合は何もしない（表示のみ）。
4. 同一結果画面の表示につき、保存ダイアログは 1 回だけ（フラグで制御）。

**CSV 内容案**（Excel で開きやすくするため BOM 付き UTF-8）:

- 1 行目: ヘッダー  
`グループID,プロジェクトID,タスクID,プロジェクト名,タスク名,問題文,結果`
- 2 行目以降: 各タスク 1 行。`結果` は `〇` / `✖` など、現在の `ResultMark` に合わせる。
- 末尾にサマリ行（例: `正答率,xx%`、`✖数,total`）を入れてもよい（任意）。

**実装**:

- [ResultWindow.xaml.cs](c:\Users\kouza\source\repos\MOS PowerPoint app\Views\ResultWindow.xaml.cs): `LoadResultsAsync` / `LoadProjectDataAsync` 完了後に、`SaveFileDialog` → CSV 出力メソッドを呼ぶ。既に保存済みかどうかのフラグを追加。
- CSV 生成は `_allProjects` とサマリ情報から組む。`ResultTaskInfo` の `ProjectId`・`TaskId`・`Description`・`ResultMark` 等を利用。
- Word 版 `ResultWindow` も同様の仕様で実装（共通化できるならヘルパーに切り出してよい）。

---

## 5. 実装順序の提案

1. **CSV 保存**
  - PowerPoint の `ResultWindow` に「表示後 1 回だけ保存ダイアログ → CSV 出力」を追加。
2. **他PC対応**
  - データパス未設定時はフォルダ選択 → `config.json` 保存、README 更新。
3. **Word 版**
  - 新規プロジェクト作成 → MainWindow / MainViewModel / UiTestAppBar / ResultWindow を Word 用に実装 → 結果画面に同じ CSV 自動保存を組み込む。

---

## 6. 注意・前提

- **Office 必須**: 他PCでも PowerPoint / Word がインストールされている必要がある。
- **データフォルダ**: 配布時は `Tab1`～`Tab3` と `Project*.pptx` / `Project*.docx` の構造をREADMEで説明し、ユーザーが用意する前提。
- **Word 問題文**: `MOS模擬アプリ問題文一覧_Word.json` は、既存の PowerPoint 用 JSON や [PP問題文.csv](c:\Users\kouza\source\repos\MOS PowerPoint app\PP問題文.csv) を参考に作成。中身は暫定でよい。

---

## 7. 変更・追加ファイル一覧（概要）


| 種別  | 対象                                                                                                        |
| --- | --------------------------------------------------------------------------------------------------------- |
| 変更  | `MainViewModel.cs`（パス選択・config 保存）、`ResultWindow.xaml` / `.cs`（CSV 保存）、`Assets\config.json`（必要に応じてスキーマ拡張） |
| 追加  | 配布用 `README.md` または `配布手順.txt`                                                                            |
| 新規  | `MOS Word app` ソリューションフォルダ一式、`MOS模擬アプリ問題文一覧_Word.json`                                                    |


以上の順で進めれば、他PC利用・Word版・結果のCSV保存の3点を満たせます。