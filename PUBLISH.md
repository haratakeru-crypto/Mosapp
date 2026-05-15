# MOSapp 発行手順（他PCで配布する場合）

他PCでも表紙（MOSapp）から各科目の exe が開き、JSON で問題文を読み込み採点できるようにするための発行手順です。

## 前提

- 発行先 1 フォルダに「表紙 exe + 3 科目 exe + 各科目の JSON/DLL/References」がすべて入っている必要があります。
- 3 科目（Excel / Word / PowerPoint）を **先に Release ビルド** してから、表紙を発行し、**必ず Merge を実行**してください。

## 方法 A: 一括スクリプト（推奨）

次の 1 コマンドで、3 科目ビルド → 表紙 Publish + VSTO → Merge まで実行します。

```powershell
powershell -ExecutionPolicy Bypass -File .\Publish-MOSapp-Full.ps1
```

- 実行場所: リポジトリルート（`Publish-MOSapp-Full.ps1` があるフォルダ）
- 発行先: `MOSapp\MOSapp\publish`（および ClickOnce 利用時は `Application Files\MOSapp_1_0_0_*` にもマージされます）

## 方法 B: 手順を分けて実行

1. **3 科目を Release ビルド**  
   Visual Studio または MSBuild で次をビルドする。
   - `MOSapp\mos_xaml_app\MOSExcelMogiApp.csproj`
   - `MOSapp\MOS Word app\MOS Word app.csproj`
   - `MOSapp\Mos PowerPoint Mogi App\MOS PowerPoint app.csproj`

2. **表紙を発行し VSTO を同梱**
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Publish-MOSapp-With-VSTO.ps1
   ```

3. **3 科目の出力を発行フォルダにマージ**
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Merge-AppOutputs-ToPublish.ps1
   ```

Merge を実行しないと、他PCで問題文・採点が参照できず動作しません。

## 発行物の確認（他PCで動かす前に）

発行フォルダ（または `Application Files\MOSapp_1_0_0_*`）に以下があることを確認してください。

- `MOSapp.exe`（表紙）
- `MOSExcelMogiApp.exe`, `MOS Word app.exe`, `Mos PowerPoint Mogi App.exe`
- `References\JSON\` 配下の各科目の問題文 JSON
- `Assets\config.json`（任意）
- 各アプリが参照する DLL（例: Newtonsoft.Json.dll など）

これらが同一フォルダにある状態で配布すると、他PCで表紙から科目を起動し、問題文の読み込みと採点が利用できます。
