---
name: チェッカーDLL置き場所変更
overview: WordChecker のソースを products\Group1 から Libraries\Group1 へ移し、DLL の出力・読み出し先を bin\products から bin\Dlls に統一する。
todos: []
isProject: false
---

# チェッカー DLL 置き場所変更プラン

## 現状と目標の対応


| 項目          | 現状                                                                   | 理想                                     |
| ----------- | -------------------------------------------------------------------- | -------------------------------------- |
| ソース         | `products\Group1\*.cs`, `*.csproj`                                   | `Libraries\Group1\*.cs`, `*.csproj`    |
| DLL 生成      | `products\Group1\bin\Debug\*.dll` + PostBuild で `bin\products\` にコピー | `bin\{Debug,Release}\Dlls\*.dll` に直接出力 |
| EXE からの読み出し | `bin\{Debug,Release}\products\*.dll`                                 | `bin\{Debug,Release}\Dlls\*.dll`       |


**注意**: リポジトリ内に `Libraries\WordChecker\Group1\` も存在しますが、メインの [MOS Word app.csproj](MOSapp/MOS Word app/MOS Word app.csproj) が参照しているのは `products\Group1\` 側のみです。本プランでは **products\Group1 を Libraries\Group1 に移す** のみ扱い、`Libraries\WordChecker\Group1` は触れません（移行後に重複があれば別途整理）。

---

## 手順

### 1. フォルダ作成とファイル移動

- **作成**: `MOS Word app\Libraries\Group1\` フォルダ（`Libraries\WordChecker\` とは別の `Libraries\Group1`）。
- **移動**（`products\Group1\` → `Libraries\Group1\`）:
  - `WordChecker1_1.cs` ～ `WordChecker1_10.cs`（10 ファイル）
  - `WordChecker1_1.csproj` ～ `WordChecker1_10.csproj`（10 ファイル）
- **移動しない**: `products\Group1\bin\`, `products\Group1\obj\` はビルド成果物のため移さず、後で削除またはクリーンで解消。

### 2. 各 WordChecker*.csproj の修正（10 プロジェクト共通）

対象: [products/Group1/WordChecker1_1.csproj](MOSapp/MOS Word app/products/Group1/WordChecker1_1.csproj) を代表とする 1_1～1_10。

- **OutputPath**: `bin\Debug\` / `bin\Release\` を `**..\..\bin\$(Configuration)\Dlls\`** に変更。  
→ DLL をメインアプリの `bin\Debug\Dlls\` / `bin\Release\Dlls\` に直接出力。
- **PostBuildEvent**: 削除（コピー不要になるため）。
- **LogReader の参照**:  
  - 現状: `<Compile Include="..\..\Libraries\LogReader.cs" Link="LogReader.cs" />`（products\Group1 からの相対パス）  
  - 変更後: `<Compile Include="..\LogReader.cs" Link="LogReader.cs" />`（Libraries\Group1 から見て `Libraries\LogReader.cs`）。

### 3. メインアプリ側の参照更新

- **[MOS Word app.csproj](MOSapp/MOS Word app/MOS Word app.csproj)**  
  - 全 `ProjectReference` のパスを `products\Group1\WordChecker1_*.csproj` → `**Libraries\Group1\WordChecker1_*.csproj`** に変更（1_1～1_10）。
- **[MainViewModel.cs](MOSapp/MOS Word app/MainViewModel.cs)**（採点時の DLL 読み込み、599～610 行付近）  
  - 第一候補: `Path.Combine(baseDir, "products", ...)` → `**Path.Combine(baseDir, "Dlls", ...)`**
  - フォールバック: `baseDir\..\products\Group1\bin\Debug\` を `**baseDir\..\Dlls\`** または `**baseDir\Dlls\`** に変更（必要なら 1 つに統一）。
  - エラーメッセージ内の「チェッカーファイルが見つかりません」のパス表記があれば `Dlls` に合わせる。

### 4. ソリューションの更新

- **[MOSapp.slnx](MOSapp/MOSapp/MOSapp.slnx)**（親ソリューション）  
  - 全 WordChecker の `Project Path`:  
  `MOS Word app/products/Group1/WordChecker1_*.csproj` → `**MOS Word app/Libraries/Group1/WordChecker1_*.csproj`**（1_1～1_10）。
- **[MOS Word app.sln](MOSapp/MOS Word app/MOS Word app.sln)**（子ソリューション、WordChecker1_1 のみ含む場合）  
  - WordChecker1_1 の相対パスを `products\Group1\WordChecker1_1.csproj` → `**Libraries\Group1\WordChecker1_1.csproj`** に変更。

### 5. 旧フォルダの扱い

- **products\Group1* は、上記変更後にビルドが通ることを確認してから削除するか、残すかを判断。  
- 削除する場合: `products\Group1\` 配下のソース・csproj のみ削除（`bin`, `obj` はクリーンまたは手動削除で可）。

---

## 変更ファイル一覧（参照のみ更新するファイル）


| ファイル                                                           | 変更内容                                                                                       |
| -------------------------------------------------------------- | ------------------------------------------------------------------------------------------ |
| 10 個の WordChecker1_*.csproj                                    | OutputPath → `..\..\bin\$(Configuration)\Dlls\`、PostBuild 削除、LogReader → `..\LogReader.cs` |
| [MOS Word app.csproj](MOSapp/MOS Word app/MOS Word app.csproj) | ProjectReference 10 件を `Libraries\Group1\`* に変更                                            |
| [MainViewModel.cs](MOSapp/MOS Word app/MainViewModel.cs)       | 読み出しパスを `products` → `Dlls`、フォールバックを `Dlls` に                                              |
| [MOSapp.slnx](MOSapp/MOSapp/MOSapp.slnx)                       | Project Path 10 件を `Libraries/Group1/`* に変更                                                |
| [MOS Word app.sln](MOSapp/MOS Word app/MOS Word app.sln)       | WordChecker1_1 のパスを `Libraries\Group1\`* に変更                                               |


---

## 実施後の確認

1. ソリューションを開き、`Libraries\Group1` 配下の 10 プロジェクトが参照されていること。
2. メインアプリをビルド（Debug / Release）し、`bin\Debug\Dlls\` / `bin\Release\Dlls\` に `WordChecker1_1.dll` ～ `WordChecker1_10.dll` が出力されること。
3. アプリ起動後、採点実行で DLL が `Dlls` から読み込まれ、採点が完了すること。
4. 問題なければ `products\Group1\` を削除（または .gitignore で無視する運用）。

