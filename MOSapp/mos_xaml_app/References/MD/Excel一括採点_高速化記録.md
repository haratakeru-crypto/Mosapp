# Excel一括採点 高速化記録（2026-05-28）

## 目的

- Excel の一括採点だけが 2 分超かかっており、Word / PowerPoint と比較して体感が悪化していたため、採点精度を維持しつつ処理時間を短縮する。

## 最終結果（確定）

- 実測: **約 45 秒**（2 回連続実行でおおむね +5 秒以内）
- 採点結果: **全問〇（満点）を確認**
- 評価: 速度と安定性のバランスが良好

## 変更方針

1. ホットパスの `AgentLog` 呼び出しオーバーヘッドを除去
2. 固定待機（`Thread.Sleep` / `Task.Delay`）を短縮
3. COM アクティブ化待機のタイムアウトを短縮
4. その後、安定性のため一部だけ値を少し戻す

## 最終採用値（ReviewPageWindow.xaml.cs）

- `ScoreAllProjects` ループ前待機: `20ms`（旧: `350ms`）
- `EnsureProjectWorkbooksReady` の再試行間待機: `80ms`（旧: `200ms`）
- `OpenExcelFilesInBackground` のブック間待機: `120ms`（旧: `500ms`）
- `TryActivateProjectWorkbook` の再試行待機:
  - 初回プロジェクト: `400ms`（調整前短縮値 `300ms` から安定性優先で微増）
  - 2 件目以降: `250ms`（調整前短縮値 `150ms` から安定性優先で微増）
- `ActivateExcelFileInternal`:
  - `timeoutMs = 900`（調整前短縮値 `700` から安定性優先で微増）
  - `pollIntervalMs = 50`（維持）
  - ウィンドウ再アクティブ化トリガー: `>300ms`

## AgentLog の扱い

- `AgentLog` は `[Conditional("ENABLE_AGENT_LOG")]` を付与し、通常ビルドでは呼び出し自体をコンパイル除去する構成に変更。
- これにより、不要な文字列生成・シリアライズ・ファイル出力のコストを避ける。

## 所感

- 最短化だけを狙うとさらに速くなる余地はあるが、再現性と誤判定リスクを考慮し、上記値を現時点の推奨値とする。
- 今後環境差（端末性能・Excel 起動状態）でばらつきが出る場合は、`TryActivateProjectWorkbook` と `timeoutMs` を優先して微調整する。
