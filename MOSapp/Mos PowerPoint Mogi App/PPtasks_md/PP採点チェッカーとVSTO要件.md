# PowerPoint 採点チェッカー・試験アプリ・VSTO 要件

## 1. PowerPointChecker の個別作成（WordChecker / ExcelChecker と同様）

- **配置先**: `C:\Users\kouza\source\repos\MOSapp\MOSapp\Mos PowerPoint Mogi App\Library\Group1`
- **形式**: WordChecker / ExcelChecker と同様に、**PowerPointChecker** として `.cs` と `.csproj` を配置する。
- **範囲**: **1_1 ～ 1_11** まで、採点チェッカーを**個別ファイル**で作成する。
- **成果物**: 各番号ごとに `.cs` と `.csproj` のデータ。

---

## 2. CSV 参照による試験アプリの作成とプロンプト参照

- **参照データ**: [PP問題文.csv](MOSapp/Mos PowerPoint Mogi App/PP問題文.csv) の CSV データを参照して試験アプリを作成する。
- **プロンプトで参照するデータ**:
  - `C:\Users\kouza\source\repos\MOSapp\MOSapp\mos_xaml_app\Libraries\Group1` の **.cs** と **.csproj** を添付し、参照する。
  - あわせて [PowerPointGrader.cs](MOSapp/Mos PowerPoint Mogi App/PowerPointGrader.cs) もプロンプトに含めて参照する。
- **うまくいかない場合**: 上記でうまくいかなかったら、undo で戻す。

---

## 3. 座標監視 VSTO（画像・プレースホルダーの誤操作を不正解とする）

- **目的**: 画像やプレースホルダーを**誤操作で動かしてしまった**場合に、正誤判定で**不正解**にできるようにする。
- **手段**: **座標を監視する VSTO** を作成する。
- **考え方**: **結果だけ**を見る採点ではなく、**過程の操作もチェック**する。

---

## 4. VSTO のログ記録方針（全問題のログ・余計な操作の監視）

- **記録範囲**: この VSTO は**すべての問題**についてログを記録する。
- **理由**: **余計な操作をしていないか**を監視するため。
- **例**:
  - 問題が「**スライドの挿入**」だけなのに、スライド内の**画像やプレースホルダーの操作**をしていた場合、
  - **結果だけ**見ると「スライドの挿入」はできているので**正解**になってしまうが、
  - **過程**を確認すると、問題文の指示に**ない操作**をしているため、**実際のテストでは不正解**とする必要がある。
