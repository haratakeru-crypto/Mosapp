# PowerPoint 採点チェッカーと試験アプリ 作成手順

## 前提

- 本手順は [PP採点チェッカーとVSTO要件.md](PP採点チェッカーとVSTO要件.md) の **##1** と **##2** の実現を目的とする。
- 試験アプリのデータソースは **MOS模擬アプリ問題文一覧_PowerPoint.json** のまま変更しない。
- 「PP問題文.csv を参照」とは、**プロンプトで AI に添付して参照させるファイル**という意味であり、試験アプリ自体が CSV を読むように変更する必要はない。

---

## ##1. PowerPointChecker の個別作成

### 目的

PowerPointGrader.cs に集約されている採点ロジックを、WordChecker / ExcelChecker と同様に個別ファイルに分割する。

### プロンプトで参照すべきファイル

##1 の作成時、以下のファイルをプロンプトに添付して参照するとよい。

| ファイル | 用途 |
|----------|------|
| [PowerPointGrader.cs](../PowerPointGrader.cs) | 分割元の採点ロジック。各 GradeProjectXTaskY を対応する PowerPointChecker1_X に移動する際の参照。 |
| [mos_xaml_app/Libraries/Group1/ExcelChecker1_1.cs](../../mos_xaml_app/Libraries/Group1/ExcelChecker1_1.cs) | クラス構成・メソッド命名規則の参考。 |
| [mos_xaml_app/Libraries/Group1/ExcelChecker1_1.csproj](../../mos_xaml_app/Libraries/Group1/ExcelChecker1_1.csproj) | 個別チェッカー用 .csproj の構造の参考。 |
| [MOS Word app/Libraries/Group1/WordChecker1_1.cs](../../MOS Word app/Libraries/Group1/WordChecker1_1.cs) | WordChecker の構成の参考（要件で WordChecker と同様と指定されているため）。 |
| [Views/UiTestAppBarWindow.xaml.cs](../Views/UiTestAppBarWindow.xaml.cs) | Task 4 の CheckSlideDeletion 等の呼び出し元。連携設計の参考。 |
| [PP問題文.csv](../PP問題文.csv) | プロジェクト・タスクの一覧。分割時の対応関係の確認用。 |

### 作成手順

1. **配置先フォルダ**
   - `Mos PowerPoint Mogi App\Libraries\Group1` に配置。

2. **1_1 ～ 1_11 の採点チェッカーを個別作成**
   - 各番号に対応する `PowerPointChecker1_X.cs` と `PowerPointChecker1_X.csproj` を作成する。
   - X は 1～11（プロジェクト1～11に対応）。

3. **PowerPointGrader.cs からの分割方針**
   - `GradeProject1Task1`～`GradeProject1Task7` → `PowerPointChecker1_1.cs`
   - `GradeProject2Task1`～`GradeProject2Task7` → `PowerPointChecker1_2.cs`
   - 以下同様に、各プロジェクトのタスク採点メソッドを対応する PowerPointChecker1_X に移動する。

4. **共通メソッドの扱い**
   - `GetSlideByNumber`, `FindShapeWithText`, `GetSlideByTitle`, `Find3DModelShape`, `FindSmartArtShape`, `IsPictureShape` 等は、各チェッカーから参照する共通モジュールに分離するか、必要なチェッカーに配置する。

5. **Task 4 の特殊処理**
   - `CheckSlideDeletion`, `ResetTask4SlideDeletionState`, `Task4PassedByThirdSlideDeletion` は、`UiTestAppBarWindow.xaml.cs` から呼び出されているため、PowerPointChecker1_1 側で利用できるように設計する。

6. **成果物**
   - 各番号ごとに `PowerPointChecker1_X.cs` と `PowerPointChecker1_X.csproj` のセット。

---

## ##2. CSV 参照による試験アプリの作成とプロンプト参照

### 目的

PP問題文.csv と採点チェッカーをプロンプトで参照させ、試験アプリの作成・拡張を AI に依頼する。

### プロンプトテンプレート

詳細な手順・ファイル一覧・プロンプト本文は [PP試験アプリ拡張_プロンプトテンプレート.md](PP試験アプリ拡張_プロンプトテンプレート.md) を参照。

### 作成手順

1. **プロンプトで参照させるファイルを準備**
   - [PP問題文.csv](../PP問題文.csv)：問題文・解答操作の参照データとして添付する。
   - [PowerPointGrader.cs](../PowerPointGrader.cs)：採点ファサード・プロジェクト・タスクの振り分け参照として添付する。
   - [Libraries\Group1\PowerPointCheckerCommon.cs](../Libraries/Group1/PowerPointCheckerCommon.cs)、[PowerPointChecker1_1.cs](../Libraries/Group1/PowerPointChecker1_1.cs) ～ [PowerPointChecker1_11.cs](../Libraries/Group1/PowerPointChecker1_11.cs)：採点ロジックの実装として添付する。
   - `mos_xaml_app\Libraries\Group1` の **.cs** と **.csproj**：ExcelChecker の構成・フォーマットの参考として添付する。

2. **試験アプリのデータソースについて**
   - 試験アプリは引き続き **MOS模擬アプリ問題文一覧_PowerPoint.json** を参照する。
   - データソースを JSON から CSV に変更する必要はない。
   - PP問題文.csv はプロンプト用の参照データである。

3. **プロンプトでの依頼内容**
   - 上記ファイルを添付したうえで、「PP問題文.csv と PowerPointGrader および PowerPointChecker を参照し、試験アプリを作成・拡張する」旨を指示する。
   - 必要に応じて、mos_xaml_app の Group1 構成（.cs / .csproj の形式）を参考にするよう指示する。

4. **うまくいかない場合**
   - undo で戻す。
