# PowerPoint 試験アプリ拡張 プロンプトテンプレート

AI に試験アプリの作成・拡張を依頼する際に、以下をプロンプトに添付して参照させる。

---

## 1. 添付するファイル一覧

| ファイル | 用途 |
|----------|------|
| [PP問題文.csv](../PP問題文.csv) | 問題文・解答操作の参照データ |
| [PowerPointGrader.cs](../PowerPointGrader.cs) | 採点ファサード・プロジェクト・タスクの振り分け参照 |
| [PowerPointCheckerCommon.cs](../Libraries/Group1/PowerPointCheckerCommon.cs) | 採点ロジックの共通ヘルパー |
| [PowerPointChecker1_1.cs](../Libraries/Group1/PowerPointChecker1_1.cs) ～ [PowerPointChecker1_11.cs](../Libraries/Group1/PowerPointChecker1_11.cs) | 採点ロジックの実装 |
| [mos_xaml_app/Libraries/Group1/ExcelChecker1_1.cs](../../mos_xaml_app/Libraries/Group1/ExcelChecker1_1.cs) および .csproj | ExcelChecker の構成・フォーマット参考 |

※試験アプリは [MOS模擬アプリ問題文一覧_PowerPoint.json](../MOS模擬アプリ問題文一覧_PowerPoint.json) を参照し続ける。PP問題文.csv は AI へのプロンプト用であり、データソースを CSV に変更する必要はない。

---

## 2. プロンプト本文（コピペ用）

```
以下のファイルを参照し、PowerPoint 試験アプリを作成・拡張してください。

- PP問題文.csv: 問題文・解答操作の参照データ
- PowerPointGrader.cs: 採点の入口（プロジェクト・タスクの振り分け）
- Libraries\Group1\PowerPointChecker*.cs: 採点ロジックの実装
- mos_xaml_app\Libraries\Group1\ExcelChecker1_1.cs および .csproj: チェッカー構成の参考

試験アプリのデータソースは MOS模擬アプリ問題文一覧_PowerPoint.json のまま使用してください。PP問題文.csv は AI への参照用であり、アプリが CSV を直接読む必要はありません。
```

---

## 3. 補足・うまくいかない場合

- 上記でうまくいかなかった場合は undo で戻す。
- 必要に応じて、mos_xaml_app の Group1 構成を参考にするようプロンプトに追記する。
