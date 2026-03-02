---
name: PP採点アプリ厳格化とVSTO
overview: PP問題文.csv を採点の手順定義として使い、未実装タスクを実装しつつ、プレースホルダー・図形の位置ずれを厳格に判定するため PowerPoint VSTO アドインでログを取得し、採点アプリでログとCOMの両方を使って正誤判定する構成にする。
todos: []
isProject: false
---

# PowerPoint 採点アプリの厳格化と VSTO による精密採点

## 現状

- **採点アプリ**: [MainViewModel.cs](MOSapp/Mos PowerPoint Mogi App/MainViewModel.cs) が [MOS模擬アプリ問題文一覧_PowerPoint.json](MOSapp/Mos PowerPoint Mogi App/MOS模擬アプリ問題文一覧_PowerPoint.json) を読み、[PowerPointGrader.cs](MOSapp/Mos PowerPoint Mogi App/PowerPointGrader.cs) で COM 経由の採点を実行している。
- **PowerPointGrader**: 11 プロジェクト中、一部タスクのみ実装（1-1～1-7, 2-1, 2-4, 3-1, 3-2, 4-4～4-6, 9-1, 9-6）。それ以外は `return false` で未実装。位置判定は 4-5 で `PositionTolerance = 2.0` を使用。
- **タスク定義**: [PP問題文.csv](MOSapp/Mos PowerPoint Mogi App/PP問題文.csv) に No, 科目, 種類, プロジェクト, タスク, 問題文, 解答操作 の 63 件が格納。採点表 [MOSPP教材用採点表.csv](MOSapp/Mos PowerPoint Mogi App/MOSPP教材用採点表.csv) は 56 問と記載（プロジェクト×タスク数で 63 に合わせる必要あり）。
- **Word 側の参考**: VSTO アドインが Ribbon コマンドを [LogReader](MOSapp/MOS Word app/Libraries/LogReader.cs) が読むパス（`%TEMP%\mos_word_log.txt`）にログ出力し、WordChecker が「ログがあれば加点」「ログ＋ドキュメント状態」で厳格に判定している。

## 方針

1. **タスク定義の正とするもの**: PP問題文.csv を正とし、アプリのタスク一覧・問題文は CSV から読み込むか、JSON を CSV と同期させる。
2. **PowerPoint VSTO アドインを新規作成**: Word と同様に、Ribbon コマンドの実行をログに記録する。加えて、図形・プレースホルダーの「位置・サイズ」が変わったタイミングでログに座標を出力し、位置ずれを厳格に判定できるようにする。
3. **採点の厳格化**: 全 63 タスクを PowerPointGrader（またはチェッカーDLL）で実装。位置・サイズが問われるタスクは「VSTO ログの座標が許容範囲内か」を優先して見る。ログが無い場合は COM のみで判定し、許容差は小さくする（例: 0.5pt）。
4. **ログの役割**: 「どのコマンドを実行したか」と「どの図形がどの座標になったか」の両方を記録し、採点時に「正しい操作ログ＋許容内の座標」で合格とする。

---

## 1. タスク定義を PP問題文.csv ベースにする

- **CSV の読み込み**: アプリ起動時または採点時に PP問題文.csv を読み、プロジェクト・タスク・問題文・解答操作を取得する。既存の JSON は、CSV から生成するか、CSV を優先して JSON はフォールバックにする。
- **採点表との対応**: MOSPP教材用採点表.csv の「56問」と PP問題文.csv の 63 行（ヘッダ除く）の差異を確認し、採点対象タスク数を 63 に統一するか、採点表の定義に合わせるかを決める（必要なら採点表を 63 タスク構成に更新）。
- **参照箇所**: [MainViewModel.ExecuteScore](MOSapp/Mos PowerPoint Mogi App/MainViewModel.cs) 内の JSON 読み込みを、CSV 読み込みに変更するか、CSV→メモリモデルに変換して同じ `project.Tasks` を渡す形にする。

---

## 2. PowerPoint VSTO アドインの新規作成

- **プロジェクト**: ソリューションに「MOS PowerPoint VSTO Add-in」を追加（Word の [New_MOSWordVSTOAddIn](MOSapp/MOS Word app/New_MOSWordVSTOAddIn/) と同構成）。PowerPoint 用の VSTO プロジェクトで Ribbon（カスタムUI）と ThisAddIn を実装。
- **ログファイル**: Word と区別するため、パスを `%TEMP%\mos_ppt_log.txt` など固定。アドインと WPF アプリの両方で同じパスを参照する（定数または設定で一元化）。
- **記録する内容**:
  - **Ribbon コマンド**: 画面切り替え・アニメーション・レイアウト・配置・書式など、採点で使いたい idMso を Ribbon.xml の `<commands>` でフックし、実行時に `[yyyy-MM-dd HH:mm:ss] [CommandId] Executed` 形式で追記。PowerPoint の idMso は Word と異なるため、実際に使うコマンドごとに検証する。
  - **図形・プレースホルダーの位置**: ユーザーが図形を移動・リサイズした結果を採点に使うため、Application の `WindowSelectionChange` や `Presentation.SlideShowNextSlide` の前後ではなく、図形のプロパティ変更を検知する方法を検討する。実装例: 一定間隔のタイマーで「選択中の Shape の Left/Top/Width/Height」を取得し、前回と変化があった場合にのみ `[timestamp] [ShapePosition] SlideIndex=n ShapeId=m Left=x Top=y Width=w Height=h` のように 1 行追記。または、Ribbon の「配置」「上揃え」等のコマンドをフックし、その直後に選択図形の座標をログに書く。
- **Ribbon**: 開発用タブ「MOSデバッグ」を追加し、ログパス表示・ログ内容表示・ログクリアボタンを用意（Word と同様）。採点アプリの「リセット」時にログクリアを呼べるように、ログファイルパスをアプリ側と共有する。

---

## 3. 採点アプリ側のログ読み込みと厳格判定

- **PP 用 LogReader**: Word 用 [LogReader](MOSapp/MOS Word app/Libraries/LogReader.cs) を参考に、PowerPoint 用の `PPLogReader`（または `LogReader` をログパスだけ切り替えて共用）を [Mos PowerPoint Mogi App](MOSapp/Mos PowerPoint Mogi App/) 内に作成。`GetLogFilePath()` は `%TEMP%\mos_ppt_log.txt`、`ReadAllLogEntries()` でコマンド行に加え、`[ShapePosition]` 行をパースして `SlideIndex, ShapeId, Left, Top, Width, Height` のリストを返す API を用意する。
- **採点前のログクリア**: プロジェクトを開いた直後や「リセット」時に、PP 用ログファイルを削除する。アドイン側の「ログクリア」ボタンと、WPF アプリ側のリセットの両方で同じファイルを消す。
- **位置ずれの厳格判定**: プレースホルダーや図形の「少しでもずれたら不正解」となるタスクでは、次の優先順位で判定する。
  1. **VSTO ログに座標が存在する場合**: 期待するスライド・図形と一致するエントリがあり、Left/Top/Width/Height が許容誤差（例: 0.5pt）以内なら合格。期待値は正解 PPT から取得するか、問題文・解答操作から決める。
  2. **ログに座標が無い場合**: 現在の PowerPointGrader と同様に COM で図形の Left/Top/Width/Height を読み、許容誤差を**厳格に**（例: 0.5pt）して判定する。現在の `PositionTolerance = 2.0` は試験的に厳しすぎる場合は 1.0 などに変更しつつ、VSTO ログがある場合はログ優先とする。

---

## 4. PowerPointGrader の拡張（全 63 タスクとログ連携）

- **未実装タスクの実装**: PP問題文.csv の 63 タスクに対応する `GradeProject*Task`* を順次実装する。既存実装（1-1～1-7, 2-1, 2-4, 3-1, 3-2, 4-4～4-6, 9-1, 9-6）を維持しつつ、2-2, 2-3, 2-5～2-7, 3-3, 3-4, 4-1～4-3, 5-1～5-5, 6-1～6-4, 7-1～7-4, 8-1～8-5, 9-2～9-5, 9-7, 10-1～10-7, 11-1～11-7 を追加する。
- **位置・サイズが関わるタスクの例とログ利用**:
  - **3-4**（スライドズームを文字より下に配置、重ならないように）: ログにスライドズーム図形の Left/Top が記録されていれば、期待範囲内かで判定。無ければ COM で Shape.Left/Top を取得し許容差を小さくして判定。
  - **4-5**（右の画像を左の画像の上端に合わせる）: 既存の COM 判定（上端の差 ≤ 許容差）を維持し、ログに「上揃え」コマンド実行＋座標が残っていればそれを優先。
  - **6-3**（3Dモデル幅 9.5、中央の枠に）: 幅 9.5 は COM で確認。中央は Left/Top のログまたは COM でスライド中央との差が許容内かで判定。
  - **11-6**（箇条書きテキストボックスを上下中央揃え）: ログの「配置」系コマンド＋図形の Top を確認、または COM で配置を取得。
- **ログを参照するタスク**: 上記のような「配置」「上揃え」「前面へ移動」等のコマンドが採点条件になるタスクでは、`PPLogReader.HasCommandExecuted("AlignTop")` のような API でログを参照し、実行されていれば COM の許容をやや緩めるか、ログ＋座標で合格とする方針を決める。

---

## 5. 補完・統一

- **リセット時のログクリア**: アプリバーやメイン画面の「リセット」で、PowerPoint のログファイル（`mos_ppt_log.txt`）を削除する処理を追加する。Word の MainViewModel の `LogReader.ClearLog()` と同様。
- **採点結果の表示**: 既存の TaskResults と ScoreResultWindow で、CSV の「問題文」「解答操作」を表示できるようにすると、受講者が何が間違いだったか把握しやすくなる。
- **VSTO アドインの配布**: アドインをインストールしていない環境ではログが空のため、位置ずれタスクは COM のみの厳格判定になる。ドキュメントに「より精密な採点には PowerPoint 用 MOS アドインのインストールが必要」と明記する。

---

## 6. 実装順序の提案

1. PP問題文.csv の読み込みとタスク一覧の CSV ベース化（MainViewModel / 問題文表示）
2. ログファイルパス定数と PPLogReader の追加（WPF アプリ内）
3. PowerPoint VSTO アドインプロジェクトの作成、Ribbon とログ出力（コマンドのみ先行）、ログパス表示・クリア
4. アドインで図形の位置ログ（選択図形の Left/Top/Width/Height を一定間隔またはコマンド実行後に記録）
5. PowerPointGrader の未実装タスクの実装（2-2, 2-3, 2-5～2-7, 3-3, 3-4, 4-1～4-3, 5-1～5-5, 6-1～6-4, 7-1～7-4, 8-1～8-5, 9-2～9-5, 9-7, 10-1～10-7, 11-1～11-7）
6. 位置・配置が関わるタスクでの PPLogReader 利用と許容誤差の厳格化（PositionTolerance の見直しとログ優先ロジック）
7. リセット時の PP ログクリアと、採点表・56/63 問の整合

---

## 7. 注意点

- **PowerPoint の idMso**: Word と異なり、Ribbon のコマンド ID が未公開・環境依存の可能性がある。アドインでフックするコマンドは、実際に PowerPoint で検証し、「不明な Office コントロール ID」が出るものは登録しない。
- **図形の同定**: ログに ShapeId を出す場合、スライドを追加/削除すると ID が変わる可能性がある。採点時は「現在開いているプレゼンテーション」のスライド・図形とログの SlideIndex/ShapeId を対応させる必要がある。
- **許容誤差**: 本番試験に近づけるため、位置は 0.5pt～1.0pt 程度にし、VSTO ログがある場合は「正しいコマンド＋ログ上の座標が許容内」で合格とする設計がよい。

