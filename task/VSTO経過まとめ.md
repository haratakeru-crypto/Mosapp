# VSTO経過まとめ

**日付**: 2025年2月24日

---

## 1. 編集記号（ShowAll）のログ

- **現状**: Word VSTOアドインでは、リボンの「編集記号の表示/非表示」（ShowAll）をクリックするたびに、`%TEMP%\mos_word_log.txt` に `[yyyy-MM-dd HH:mm:ss] [ShowAll] Executed` 形式で記録されている。
- **実装**: `Ribbon.xml` の `<command idMso="ShowAll" onAction="CommandOnAction" />` によりフックし、`Ribbon.CommandOnAction` → `Logger.LogCommand(commandId)` でログ出力。
- **方針**: 編集記号のログを優先し、文字の入力・削除のログは行わない旨を `Logger.cs` および `Ribbon.cs` のコメントで明示。

---

## 2. 1-1 の採点でログを利用

- **要件**: 1-1（編集記号の表示/非表示）は「ログを優先し、ログがなければ不正解」。  
  「2回以上クリックし、かつ最終的に編集記号が表示」の場合のみ正解とする。
- **変更**: `products/Group1/WordChecker1_1.cs` の `CheckTask_1_1_01` を次のように変更。
  - ログファイルが存在しない → 不正解
  - ログ内で ShowAll の実行が 2 回未満 → 不正解
  - 上記を満たし、かつ `View.ShowAll == true`（編集記号が表示）→ 正解

---

## 3. 初期状態で誤って正解になる問題

- **現象**: 編集記号を一度も触っていない（初期のONのまま）でも、1-1 が正解になってしまう。
- **原因**: ログファイル `mos_word_log.txt` が %TEMP% に残り続けるため、過去のセッションや別のテストで記録された ShowAll がそのまま残り、採点時に「2回以上」と判定されていた。
- **対応**: 受験開始時（リセット時）にログをクリアするようにし、その回の操作だけを採点に反映する。

---

## 4. ログクリアのタイミング

- **条件**: 次の2つの操作のときにログをクリアする。
  1. **プロジェクト一覧画面の「すべてリセット」ボタン**を押したとき
  2. **アプリバーの「リセット」ボタン**を押したとき

- **実装**:
  - `Libraries/LogReader.cs`: ログファイルを削除する **`ClearLog()`** を追加。
  - `MainViewModel.cs`: 「すべてリセット」でユーザーが「はい」を選択した直後に `LogReader.ClearLog()` を実行。
  - `Views/AppBarWindow.xaml.cs`: 「リセット」でユーザーが「はい」を選択した直後、`ResetProject` の前に `LogReader.ClearLog()` を実行。

---

## 5. 確認・運用のポイント

- **VSTOの確認**: スタートアップを `MOSWordVSTOAddIn` にして F5 で Word を起動し、編集記号をクリックして `%TEMP%\mos_word_log.txt` に `[ShowAll] Executed` が追記されることを確認（詳細は `VSTO_デバッグ手順.md` 参照）。
- **採点**: リセット後に Word で「編集記号を非表示→再度表示」を実行し、その状態で採点すると 1-1 が正解になる。リセットせずに前回のログが残ったまま採点すると、初期状態でも正解になるため、受験開始前のリセットが必須。

---

## 参照ファイル

| 役割           | ファイル |
|----------------|----------|
| ログ読み・クリア | `Libraries/LogReader.cs` |
| 1-1 採点       | `products/Group1/WordChecker1_1.cs` |
| VSTO ログ出力  | `MOSWordVSTOAddIn/Logger.cs`, `Ribbon.xml`, `Ribbon.cs` |
| リセット時クリア | `MainViewModel.cs`, `Views/AppBarWindow.xaml.cs` |
