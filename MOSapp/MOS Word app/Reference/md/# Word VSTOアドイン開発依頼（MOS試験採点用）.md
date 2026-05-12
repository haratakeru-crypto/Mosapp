# Word VSTOアドイン開発依頼（MOS試験採点用）

あなたはC#とVSTOのエキスパートです。
Wordの操作ログを記録し、MOS模擬試験の採点（プロセス評価）を行うためのアドインを開発しています。

提供された要件に基づき、特定のコマンド実行をフックしてログに出力する「Ribbon XML」と「C#コード」を作成してください。

## 1. 開発の目的
ユーザーが「正解の手順」で操作したかを判定するため、特定のボタンやコマンドが実行された瞬間に、その操作IDとタイムスタンプをログファイルに記録します。

## 2. 監視対象の操作リスト（要件）
以下のプロジェクト・タスクに対応する操作を監視対象として実装してください。

| ID | 操作内容 | 監視すべきコマンド (idMso想定) | 備考 |
|:---|:---|:---|:---|
| **1-1** | 編集記号の表示/非表示 | `ShowAll` | トグルボタン。On/Offに関わらずクリックを検知 |
| **2-1** | 文字列の切り取り | `Cut` | リボン上の「切り取り」ボタン |
| **6-1** | 文字列を表にする | `TableConvertTextToTable` | [挿入]タブ > [表] > [文字列を表にする] |
| **8-1** | 自動作成の目次2を挿入 | `TocAutomatic2` | [参考資料] > [目次] ギャラリー内の特定アイテム |

## 3. 実装要件

### A. Ribbon XML (`Ribbon.xml`)
- 上記の `idMso` をターゲットにした `<commands>` タグ定義を行ってください。
- 既存のUIを隠さず、裏側でイベントだけをフックすること。

### B. コールバック処理 (`Ribbon.cs` / `ThisAddIn.cs`)
- `onAction` イベントハンドラを実装してください。
- 実行されたコマンドのID（例: "Cut", "ShowAll"）を識別し、ログ出力メソッドに渡してください。
- **重要:** ログ出力後、`cancelDefault = false` とし、Word本来の機能が正常に動作するようにしてください。

### C. ログ出力仕様
- 出力先: `Path.GetTempPath()` 内の `mos_word_log.txt`
- 出力形式: `[yyyy-MM-dd HH:mm:ss] [CommandID] Executed`
- 排他制御: `lock` を使用してファイル書き込みの競合を防ぐこと。

## 4. 特記事項
- **目次（Gallery）の扱い:** `TocAutomatic2` はギャラリー内のアイテムですが、もし特定のアイテムIDのフックが難しい場合、親の `TableOfContentsGallery` をフックして、引数から選択されたIDを取得する方法、または代替案を提案してください。

## 出力してほしいコード
1. `Ribbon.xml` の全量
2. `Ribbon.cs` の `onAction` メソッドおよび必要なヘルパーメソッド
3. ログ書き込みクラス（Logger）

---

## プロジェクト構造と配置先

```
MOS Word app/
└── MOSWordVSTOAddIn/                 # VSTOアドイン専用プロジェクト
    ├── MOSWordVSTOAddIn.sln          # ソリューション
    ├── MOSWordVSTOAddIn.csproj       # プロジェクトファイル
    ├── Logger.cs                     # ログ出力クラス（ルート）
    ├── Ribbon.cs                     # Ribbonコールバック（ルート）
    ├── Ribbon.xml                    # Ribbon定義（ルート、埋め込みリソース）
    ├── ThisAddIn.cs                  # アドインエントリーポイント（ルート）
    ├── ThisAddIn.Designer.cs         # 自動生成補助（ルート）
    ├── Properties/
    │   └── AssemblyInfo.cs
    ├── packages.config
    └── bin/Debug/                    # ビルド出力（.dll, .manifest, .vsto）
```

### 重要な設定
- `Ribbon.xml` のビルドアクション: `埋め込みリソース (Embedded Resource)`
- `ToolsVersion`: `Current`
- `<ProjectTypeGuids>`: `{BAA0C2D2-18E2-41B9-852F-F413020CAA33};{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}`
- `<OfficeApplication>Word</OfficeApplication>`, `<VSTOProjectType>AddIn</VSTOProjectType>`
- ServiceGuid は csproj のメイン PropertyGroup 内に配置

---

## Visual Studio での VSTO プロジェクト作成手順

1. **前提ワークロードを確認**
   - Visual Studio インストーラーで「Office/SharePoint 開発」ワークロードを有効にする。
2. **新規プロジェクト作成**
   - テンプレート: 「Word 2013 および 2016 VSTO アドイン」(または Word VSTO Add-in)
   - プロジェクト名: `MOSWordVSTOAddIn`
3. **プロパティ設定 (csproj)**
   - Framework: `.NET Framework 4.8`
   - ToolsVersion: `Current`
   - ProjectTypeGuids: 上記を設定
   - OfficeApplication: `Word`
   - VSTOProjectType: `AddIn`
4. **ファイル配置**
   - `Ribbon.xml`, `Ribbon.cs`, `Logger.cs`, `ThisAddIn.cs` をプロジェクト直下に配置
   - `Ribbon.xml` は埋め込みリソースに設定
5. **ビルド & 出力確認**
   - Debug ビルド後、`bin/Debug/` に `.dll`, `.dll.manifest`, `.vsto` が生成されることを確認

---

## 実行可能なプロンプト（AI向け）

以下をそのまま渡すと、アドイン一式（Ribbon.xml, Ribbon.cs, Logger.cs, ThisAddIn.cs）を生成できます。

```
あなたはC#とVSTOのエキスパートです。以下仕様でWord VSTOアドインを作成してください。

[目的]
- Word標準コマンドをフックし、バックグラウンドで操作ログを %TEMP%/mos_word_log.txt に記録する。
- 主用途: MOS模擬試験のプロセス評価。

[監視対象 idMso と注意点]
- ShowAll (1-1 編集記号の表示/非表示) トグルだがクリックを検知
- Cut (2-1 文字列の切り取り) ショートカット(Ctrl+X)も検知
- TableConvertTextToTable (6-1 文字列を表にする) ダイアログ直前を検知
- TableOfContentsGallery (8-1 目次ギャラリー) 個別アイテムはXMLで直接取得不可。親ギャラリーを監視

[プロジェクト構成と配置]
- ソリューション: MOSWordVSTOAddIn/MOSWordVSTOAddIn.sln
- プロジェクト: MOSWordVSTOAddIn/MOSWordVSTOAddIn.csproj
- ルート直下に配置: Ribbon.xml (埋め込みリソース), Ribbon.cs, Logger.cs, ThisAddIn.cs, ThisAddIn.Designer.cs
- プロパティ: ToolsVersion=Current, TargetFrameworkVersion=v4.8
- ProjectTypeGuids={BAA0C2D2-18E2-41B9-852F-F413020CAA33};{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}
- OfficeApplication=Word, VSTOProjectType=AddIn
- ServiceGuid は csproj のメイン PropertyGroup に設定

[実装要件]
1) Ribbon.xml
- commands タグで上記 idMso を onAction に割り当て
- BuildAction: Embedded Resource

2) Ribbon.cs
- IRibbonExtensibility を実装
- GetCustomUI で Ribbon.xml を返す
- onAction: CommandOnAction (または RepurposeOnAction) で control.Id を Logger に渡す
- TableOfContentsGallery は onAction で selectedId を受け取り、TocAutomatic2 などをログに書く（取得不可の場合はギャラリー操作を記録）
- cancelDefault = false で Word 本来の動作を継続

3) Logger.cs
- 出力: %TEMP%/mos_word_log.txt
- 形式: [yyyy-MM-dd HH:mm:ss] [CommandID] Executed (必要なら extraInfo も付加)
- UTF-8 Append、lock で排他制御、例外は握りつぶしてWordに影響させない

4) ThisAddIn.cs
- CreateRibbonExtensibilityObject で Ribbon インスタンスを返す
- Startup/Shutdown を定義（デバッグログ程度で可）

[成果物]
- Ribbon.xml, Ribbon.cs, Logger.cs, ThisAddIn.cs を生成し、ビルド可能な状態にすること。
```

---

## 備考（目次ギャラリーの制約と採点のヒント）
- Ribbon XML では TocAutomatic2 の個別アイテム ID を直接フックできないため、TableOfContentsGallery 全体を監視し、ログには「目次機能が使われた」事実を残す。
- 採点時は生成されたドキュメントの目次タイトル等で自動作成1/2 を判定する（例: 自動作成1=Contents、自動作成2=Table of Contents）。