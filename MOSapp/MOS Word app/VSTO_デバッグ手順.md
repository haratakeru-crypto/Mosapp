# VSTOアドインデバッグ手順

## 問題の原因

現在、WPFアプリケーション（MOS Word app.exe）をデバッグしていますが、VSTOアドインは**Wordプロセス内で動作する別のプロセス**です。そのため、VSTOアドインをデバッグするには、Wordを起動してアドインを読み込む必要があります。

## デバッグ手順

### 1. Visual Studioでソリューションを開く

1. Visual Studioで`MOS Word app.sln`を開きます

### 2. VSTOプロジェクトをスタートアッププロジェクトに設定

1. ソリューションエクスプローラーで`MOSWordVSTOAddIn`プロジェクトを右クリック
2. 「スタートアッププロジェクトに設定」を選択

### 3. ビルドとデバッグ

1. **F5キー**を押すか、「デバッグ」→「デバッグの開始」をクリック
2. Wordが自動的に起動します
3. VSTOアドインがWordに読み込まれます
4. デバッグログに`[MOSWordVSTOAddIn] Add-in started`が表示されることを確認

### 4. ログの確認

VSTOアドインが正常に動作している場合、以下のログが出力されます：

```
[MOSWordVSTOAddIn] Add-in started
[MOSWordVSTOAddIn] Log file: C:\Users\kouza\AppData\Local\Temp\mos_word_log.txt
```

### 5. 動作確認

1. Wordで何か操作を実行（例：編集記号の表示/非表示）
2. `%TEMP%\mos_word_log.txt`にログが記録されることを確認

## トラブルシューティング

### VSTOアドインが読み込まれない場合

1. **ビルドエラーの確認**
   - ソリューションエクスプローラーで`MOSWordVSTOAddIn`プロジェクトを右クリック
   - 「ビルド」を選択してエラーがないか確認

2. **Wordのアドイン管理を確認**
   - Wordを起動
   - 「ファイル」→「オプション」→「アドイン」
   - 「管理」で「COMアドイン」を選択して「設定」をクリック
   - `MOSWordVSTOAddIn`が表示されているか確認

3. **レジストリの確認**
   - レジストリエディタを開く
   - `HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\MOSWordVSTOAddIn`を確認
   - `LoadBehavior`が`3`（起動時に読み込む）になっているか確認

### ビルドエラーが発生する場合

1. **Visual StudioのOffice開発ツールがインストールされているか確認**
   - Visual Studioインストーラーを開く
   - 「Office/SharePoint開発」ワークロードがインストールされているか確認

2. **プロジェクトファイルの確認**
   - `MOSWordVSTOAddIn.csproj`が正しく読み込まれているか確認
   - Visual Studioでプロジェクトを再読み込み

## 採点前に確認すること（Dlls の配置）

採点はメインアプリが `bin\Debug\Dlls\`（または `bin\Release\Dlls\`）内の `WordChecker1_1.dll` ～ `WordChecker1_10.dll` を読み込んで実行します。

1. ソリューションをビルドする
2. メインアプリの実行フォルダ（例: `MOS Word app\bin\Debug`）の下に `Dlls` フォルダがあることを確認
3. `Dlls` 内に `WordChecker1_1.dll` ～ `WordChecker1_10.dll` が存在することを確認
4. 不足している場合は、`Libraries\Group1` の各 WordChecker プロジェクトのビルド出力が `..\..\bin\$(Configuration)\Dlls\` になっているか確認し、ソリューション全体をリビルドする

## New_MOSWordVSTOAddIn でデバッグする場合

- スタートアッププロジェクトを **New_MOSWordVSTOAddIn** に設定して F5 で Word を起動する
- 出力ウィンドウで `[Ribbon] GetCustomUI called` および `[Ribbon] Ribbon_Load completed` が表示されれば、リボン（MOSデバッグタブ）とコマンドフックが有効です
- リボンが表示されない場合は、一度アドインをアンインストールし、最新ビルドの `.vsto` で再インストールしてから Word を起動し直す

## 注意事項

- VSTOアドインとWPFアプリケーションは**別プロセス**で動作します
- VSTOアドインのログは`%TEMP%\mos_word_log.txt`に出力されます
- WPFアプリケーションは`Libraries\LogReader.cs`を使用してログを読み込みます




