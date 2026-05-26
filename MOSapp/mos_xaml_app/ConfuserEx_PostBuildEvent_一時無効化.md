# ConfuserEx PostBuildEvent 一時無効化メモ

## 背景

- `MOSExcelMogiApp.csproj` の `PostBuildEvent` で、以下のコマンドが実行される構成だった。
- `"C:\apps\ConfuserEx-CLI\Confuser.CLI.exe" -n "$(ProjectDir)$(ProjectName).crproj"`
- ローカル環境に `Confuser.CLI.exe` が存在しない場合、`MSB3073` でビルドが失敗する。

## 今回の対応（2026-03-27）

- `Debug` 実行を優先するため、`MOSExcelMogiApp.csproj` から `PostBuildEvent` を一時的に削除。
- この対応は、アプリ本体の実行確認と開発継続を目的とした暫定措置。

## 復活させるとき

以下を `MOSExcelMogiApp.csproj` の末尾（`</Project>` の直前）に戻す。

```xml
<PropertyGroup>
  <PostBuildEvent>"C:\apps\ConfuserEx-CLI\Confuser.CLI.exe" -n "$(ProjectDir)$(ProjectName).crproj"</PostBuildEvent>
</PropertyGroup>
```

## 推奨（再発防止）

- 将来的には `Release` のみ難読化を実行し、`Debug` では実行しない条件付き設定にする。
- 例: `Condition="'$(Configuration)'=='Release'"` を使って `PostBuildEvent` または `Target` を分岐。
