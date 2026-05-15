---
name: VSTOアドインの完成と設定
overview: 既存のVSTOアドインファイルを完成させ、Visual Studioでのビルドと配置が可能な状態にする
todos:
  - id: vsto-1
    content: MOSWordVSTOAddIn.csprojにVSTOビルドターゲットのインポートを追加
    status: pending
  - id: vsto-2
    content: プロジェクトファイルの構成を確認し、必要に応じて修正
    status: pending
    dependencies:
      - vsto-1
  - id: vsto-3
    content: README.mdに手動作業の手順を追加・更新
    status: pending
    dependencies:
      - vsto-1
      - vsto-2
---

# VST

Oアドインの完成と設定

## 現状

既に基本的なVSTOアドインのファイルは作成されていますが、Visual Studioでのビルドと配置を成功させるために、いくつかの修正と手動作業が必要です。

## 実装内容

### 1. プロジェクトファイルの修正

#### MOSWordVSTOAddIn.csproj

- VSTOのビルドターゲットをインポートするImport要素を追加

- 必要に応じてVSTO固有のプロパティを追加

#### 修正箇所

`MOSWordVSTOAddIn/MOSWordVSTOAddIn.csproj`の最後に、`<Import Project="$(MSBuildToolsPath)\Microsoft.CSharp.targets" />`の前に、以下のインポートを追加：

```xml
<Import Project="$(MSBuildExtensionsPath32)\Microsoft\VisualStudio\v$(VisualStudioVersion)\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets" Condition="'$(VisualStudioVersion)' != '' And '$(VisualStudioVersion)' != '10.0' And '$(VisualStudioVersion)' != '11.0'" />
```



### 2. ThisAddIn.csの確認と修正

- `ThisAddIn`クラスは`Microsoft.Office.Tools.Word.AddInBase`を継承している必要があるが、VSTOが自動生成する部分クラスと結合される

- 現在の実装は問題ないが、VSTOが生成するコードと結合するために、プロジェクトが正しくVSTOテンプレートとして認識される必要がある

### 3. リソースファイルの設定確認

- `Ribbon.xml`がEmbeddedResourceとして正しく設定されていることを確認

- 現在の`.csproj`では既に設定済み

### 4. 参照の確認

- Microsoft.Office.Tools.Word.v4.0.Utilities

- Microsoft.VisualStudio.Tools.Applications.Runtime

- これらの参照が正しく設定されていることを確認（現在設定済み）

## 手動作業（Visual Studioでの作業）

### ステップ1: Visual Studioでプロジェクトを開く

1. Visual Studioでソリューション（`MOS Word app.sln`）を開く

2. `MOSWordVSTOAddIn`プロジェクトが正しく読み込まれていることを確認

### ステップ2: VSTOプロジェクトの設定確認

1. `MOSWordVSTOAddIn`プロジェクトを右クリック > [プロパティ]

2. [アプリケーション]タブで以下を確認：

- [アセンブリ名]: `MOSWordVSTOAddIn`

- [ルート名前空間]: `MOSWordVSTOAddIn`

3. [ビルド]タブで以下を確認：

- [プラットフォームターゲット]: `Any CPU`または`x86`（Officeのバージョンに応じて）

4. [発行]タブ（または[セキュリティ]タブ）で以下を確認：

- 証明書の設定（開発用は自動生成される）

### ステップ3: 必要なNuGetパッケージの復元

1. ソリューションエクスプローラーで`MOSWordVSTOAddIn`プロジェクトを右クリック

2. [NuGetパッケージの復元]を実行（必要な場合）

### ステップ4: ビルドとテスト

1. `MOSWordVSTOAddIn`プロジェクトをビルド（Ctrl+Shift+B）

2. エラーがないことを確認

3. 出力ディレクトリ（`bin\Debug\`または`bin\Release\`）に以下が生成されることを確認：

- `MOSWordVSTOAddIn.dll`

- `MOSWordVSTOAddIn.dll.manifest`（VSTOが自動生成）

- `MOSWordVSTOAddIn.vsto`（VSTOが自動生成）

### ステップ5: アドインの登録とテスト

1. Wordを終了

2. Visual Studioで`MOSWordVSTOAddIn`プロジェクトを右クリック > [デバッグ] > [新しいインスタンスを開始]

- または、F5キーでデバッグ開始

3. Wordが起動し、アドインが読み込まれることを確認

4. [ファイル] > [オプション] > [アドイン] > [COMアドイン]で、`MOSWordVSTOAddIn`が有効になっていることを確認

### ステップ6: ログ機能のテスト

1. Wordで以下の操作を実行：

- [ホーム]タブ > [段落] > [編集記号の表示/非表示]（ShowAll）

- テキストを選択して切り取り（Cut）
- 貼り付け（Paste）

2. `%TEMP%\mos_word_log.txt`を開いて、ログが記録されていることを確認

### ステップ7: 手動登録（デバッグが動作しない場合）

1. `bin\Debug\MOSWordVSTOAddIn.vsto`を右クリック > [インストール]

2. または、以下のコマンドを実行：

   ```javascript
      "%ProgramFiles(x86)%\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\gacutil.exe" /i "MOSWordVSTOAddIn.dll"
   ```

3. Wordを起動して、アドインが読み込まれることを確認

## トラブルシューティング

### エラー: "VSTOプロジェクトとして認識されない"

- Visual Studioの「Office/SharePoint開発」ワークロードがインストールされていることを確認

- プロジェクトの`ProjectTypeGuids`が`{BAA0C2D2-18E2-41B9-852F-F413020CAA33}`を含んでいることを確認（現在設定済み）

### エラー: "参照が見つからない"

- `Microsoft.Office.Tools.Word`などの参照パスを確認

- Officeのバージョンに応じて、参照のバージョンを調整

### エラー: "アドインが読み込まれない"

- Wordのセキュリティ設定を確認（[ファイル] > [オプション] > [セキュリティセンター] > [セキュリティセンターの設定] > [アドイン]）

- 証明書を信頼済みに追加

### エラー: "Ribbon.xmlが見つからない"

- `Ribbon.xml`がEmbeddedResourceとして正しく設定されていることを確認

- ビルドアクションが「埋め込みリソース」になっていることを確認

## 実装ファイル

### 修正するファイル

- `MOSWordVSTOAddIn/MOSWordVSTOAddIn.csproj`: VSTOビルドターゲットのインポートを追加

### 確認するファイル（変更不要）

- `MOSWordVSTOAddIn/ThisAddIn.cs`: 既に正しく実装されている

- `MOSWordVSTOAddIn/Ribbon.cs`: 既に正しく実装されている

- `MOSWordVSTOAddIn/Logger.cs`: 既に正しく実装されている

- `MOSWordVSTOAddIn/Ribbon.xml`: 既に正しく実装されている

- `MOSWordVSTOAddIn/Properties/AssemblyInfo.cs`: 確認のみ

## 次のステップ