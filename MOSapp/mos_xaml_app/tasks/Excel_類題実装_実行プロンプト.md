# Excel 類題作成 — Cursor 実行プロンプト集

このファイルは、**Cursor に依頼するときにコピペするだけ**で使えるプロンプト集です。

**依頼の目的**

- 類題の Excel・Checker・問題文を作る
- 採点が動くところまで VS で確認する

**依頼に含めないもの**

- アプリバーの「類題ボタン」
- 類題モードの自動切替（`GetProjectFilePath` 改修、`practiceVariants` 追加など）
- `AppBarWindow.xaml` の UI 改修

仕様の詳細は、先にこちらを読ませてください。

- `tasks/Excel_類題作成_実装土台ガイド.md`

---

## 0. 使い方 — どのプロンプトを使うか

```
初めて試すとき
  → §1「先行体験プロンプト」を貼る（タスク1だけ採点・問題文）
       ↓ 成功したら
  → §2「本番プロンプト」を貼る（1プロジェクト全体を仕上げる）
```

1. Cursor を **Agent モード** にする
2. **初回は §1 先行体験** → うまくいったら **§2 本番**
3. Checker 名は土台ガイド **§3 の表** を参照（プロジェクト番号と Checker 番号は一致しない）

**Checker の作り方:** Terminal でファイルを複製（`copy` / `cp`）しない。  
教材 Checker を参照し、**新規 .cs ファイル** として類題 Checker を作成する。

---

## 1. 先行体験プロンプト（最初にこれを貼る）

採点と問題文が **タスク1だけ** 動く成功体験用です。残りのタスク・問題文は **§2 本番** で仕上げます。

```text
@MOSapp/mos_xaml_app/tasks/Excel_類題作成_実装土台ガイド.md
@MOSapp/mos_xaml_app/tasks/Excel_類題実装_実行プロンプト.md

土台ガイド §5「先行体験」に従い、採点と問題文をタスク1だけ試せる状態を作ってください。

【今回の対象】
- 画面上プロジェクト1
- 類題1（PV1）
- タスク1のみ（採点: CheckTask_1_2_01 / 問題文: projectId:1 の taskId:1）

【類題の内容（ここに書く）】
- タスク1の類題問題文:
  （例: 「シート［試験結果］のテーブルに、見出し1スタイルを設定します。」）
- タスク1の正解条件の要点:
  （例: テーブル縞模様を列方向に設定、など Checker に反映する条件）

※ 上記が空でも、教材タスク1を参考に類題案を1つ提案して進めてよい。

【やること】
1. 教材 ExcelChecker1_2.cs を参照し、ExcelChecker1_2_PV1.cs を新規作成（Terminal 複製しない）
2. CheckTask_1_2_01 だけ類題の正解条件に合わせて実装
3. CheckTask_1_2_02〜05 は教材 ExcelChecker1_2 の同名メソッドを呼び出して委譲
4. MOS演習問題文一覧.json の projectId:1 / taskId:1 の description だけ類題文言に更新
   （他の taskId は変更しない）
5. mos_xaml_app/Assets/config.json の project "1" の library を ExcelChecker1_2_PV1 に変更
6. Debug ビルド

【今回やらないこと】
- タスク2以降の類題ロジック実装（委譲のみ）
- 問題文 JSON の taskId:2 以降の変更
- Excel ファイルの編集（教材担当が別途行う想定。未編集でもビルドは通す）
- 類題ボタン・GetProjectFilePath 改修・practiceVariants 追加
- ランチャー config・Word/PowerPoint の変更

【Checker 作成ルール】
- Terminal で copy/cp によるファイル複製は使わない
- 教材 Checker は読み取り参照のみ（上書きしない）

【完了報告（必須）】
1. 参照した教材 Checker
2. 新規作成した ExcelChecker1_2_PV1.cs
3. タスク1だけ類題化したこと、タスク2以降は委譲したこと
4. 問題文 JSON の変更（taskId:1 の変更前後を1行ずつ）
5. config.json の library 変更内容
6. VS での確認手順（土台ガイド §5.5 の要約）
7. 次に §2 本番プロンプトへ進むべきか
```

---

## 2. 本番プロンプト（先行体験のあとに貼る）

1プロジェクト分を一通り仕上げるときに使います。**§1 が終わってから** 貼ってください。

```text
@MOSapp/mos_xaml_app/tasks/Excel_類題作成_実装土台ガイド.md
@MOSapp/mos_xaml_app/tasks/Excel_類題実装_実行プロンプト.md

土台ガイド §7「本番作業」に従い、Excel 演習タブ（Tab1）の類題を1プロジェクト分仕上げてください。
§1 先行体験で ExcelChecker1_2_PV1.cs がある場合はそれを拡張し、作り直さないでください。

【今回の対象】
- 類題1（PracticeVariant1 / _PV1）
- プロジェクト1 だけ（全10プロジェクトは今回やらない）

【やること（土台ガイド §7 準拠）】
1. Tab1\Initial\project1.xlsx を類題用に編集（Templates も必要なら）
2. §3.1 の表に従い、ExcelChecker1_2_PV1.cs の全タスクを類題用に実装（委譲をやめて本体に書く）
   （参照元教材 Checker: ExcelChecker1_2。config の "1" を参照）
3. MOS演習問題文一覧.json の projectId:1 部分を類題文言に更新
   （原稿用に MOS演習問題文一覧_PracticeVariant1.json も新規作成してよい）
4. mos_xaml_app/Assets/config.json の project "1" の library を ExcelChecker1_2_PV1 に変更（未変更なら確認のみ）
5. Debug ビルドを実行し、エラーがあれば修正

【Checker 作成ルール】
- Terminal で copy/cp によるファイル複製は使わない
- 教材 Checker ファイルは読み取り参照のみ（上書きしない）
- 類題 Checker は Libraries/Group1/ に新規 .cs として追加（または既存 _PV1 を拡張）
- 完了報告に「参照した教材 Checker」「新規作成したファイル」を必ず書く

【やらないこと】
- Terminal で Checker ファイルを複製しない
- AppBar に類題ボタンを付けない
- GetProjectFilePath / LoadTasksFromJson / MainViewModel のモード切替を実装しない
- practiceVariants 設定を config.json に追加しない
- MOSapp/MOSapp/Assets/config.json（ランチャー）を触らない
- Word / PowerPoint アプリを触らない
- 破壊的操作（ExcelTaskValidationConfig.cs）は、必要がなければ触らない

【命名・対応】
- プロジェクト番号と Checker 番号は一致しない（§3 参照）
- Checker: ExcelChecker1_2_PV1（必ず ExcelChecker で始める）
- メソッド: CheckTask_1_2_01 形式を維持

【ビルド】
- mos_xaml_app で Debug ビルド
- コマンドは .cursor/rules/build-method.mdc を参照

【完了報告（必須）】
1. 参照した教材 Checker（例: ExcelChecker1_2.cs）
2. 新規作成した類題 Checker（例: ExcelChecker1_2_PV1.cs）
3. その他変更・追加したファイル一覧
4. config.json で library をどう変えたか
5. VS 確認手順（§7 Step 5 の要約）
6. §8 チェックリストの達成状況
7. 手動確認が必要な項目（あれば）
```

---

## 3. 追加オプション（必要なときだけ足す）

§1 または §2 のプロンプト末尾に追加してください。

### 3-A. Checker だけ作りたい（本番用）

```text
【追加指示】
Excel と JSON は触らない。ExcelChecker1_2_PV1.cs を新規作成してビルドのみ。
Terminal でのファイル複製は使わない。config.json の library 変更も今回はしない。
```

### 3-B. 問題文 JSON だけ作りたい

```text
【追加指示】
Checker と Excel は触らない。
MOS演習問題文一覧_PracticeVariant1.json を作成し、
projectId:1 の description を類題用に書く。メイン JSON への反映も行う。
```

### 3-C. 1プロジェクトずつ進めたい

```text
【追加指示】
今回は project1 のみ。他プロジェクトには着手しない。
完了報告後に止まる。
```

### 3-D. 次のプロジェクト（例: プロジェクト2）に進みたい

```text
【追加指示】
土台ガイド §3.1 の表に従い、プロジェクト2 で同じ作業を行う。
- Excel: project2.xlsx（中身は元プロジェクト3）
- 参照元 Checker: ExcelChecker1_3（新規ファイル ExcelChecker1_3_PV1.cs を作成）
- 類題 Checker: ExcelChecker1_3_PV1
- config の project "2" の library を ExcelChecker1_3_PV1 に変更
```

### 3-E. プロジェクト7（ずれあり）に進みたい

```text
【追加指示】
土台ガイド §3.3 の表に従い、プロジェクト7 で同じ作業を行う。
- Excel: project7.xlsx（中身は元プロジェクト9）
- 参照元 Checker: ExcelChecker1_9（※ 7 ではない。新規ファイル ExcelChecker1_9_PV1.cs を作成）
- 類題 Checker: ExcelChecker1_9_PV1
- config の project "7" の library を ExcelChecker1_9_PV1 に変更
```

---

## 4. 作業ステップ別プロンプト（1つずつ依頼するとき）

土台ガイド §7 の Step に対応しています（本番作業）。先行体験は §1 を使ってください。

### Step 2 — Checker 作成（本番）

```text
@MOSapp/mos_xaml_app/tasks/Excel_類題作成_実装土台ガイド.md §7 Step 2

project1 向けの類題 Checker ExcelChecker1_2_PV1.cs を新規作成してください（§3.1 参照）。
教材の ExcelChecker1_2.cs を参照し、類題の正解条件に合わせて内容を書く。
Terminal でのファイル複製は使わない。ビルドして報告。config.json と UI は触らない。
完了報告に参照元・新規作成ファイルを明記すること。
```

### Step 3 — config の library 変更

```text
@MOSapp/mos_xaml_app/tasks/Excel_類題作成_実装土台ガイド.md §7 Step 3

mos_xaml_app/Assets/config.json の project "1" の library を ExcelChecker1_2_PV1 に変更。
他のプロジェクトやランチャー config は触らない。
```

### Step 4 — 問題文 JSON

```text
@MOSapp/mos_xaml_app/tasks/Excel_類題作成_実装土台ガイド.md §7 Step 4

projectId:1 の類題用問題文を JSON に反映してください。
MOS演習問題文一覧.json を更新。原稿用に _PracticeVariant1 版も作成可。
```

### Step 5 — ビルド確認

```text
@MOSapp/mos_xaml_app/tasks/Excel_類題作成_実装土台ガイド.md §7 Step 5

Debug ビルドを実行し、§8 チェックリストに沿って確認可能な状態か報告してください。
アプリの UI 改修はしない。
```

---

## 5. 続き依頼用プロンプト

```text
@MOSapp/mos_xaml_app/tasks/Excel_類題作成_実装土台ガイド.md
@MOSapp/mos_xaml_app/tasks/Excel_類題実装_実行プロンプト.md

前回の類題作成作業の続きをしてください。
- 前回の完了報告と変更ファイルを確認してから着手
- 類題ボタンやアプリ改修はしない
- 未完了の Step のみ実施（勝手に範囲を広げない）
- 完了後は §1 または §2 の【完了報告】形式で報告
```

---

## 6. うまくいかないときのプロンプト

```text
類題の作成・採点確認で問題が出ています。調査・修正してください。

【症状】
（例: 採点しても全部 ✖ になる）

【期待】
（例: 正しい操作後は全部 〇）

【制約】
- 土台ガイドと実行プロンプトに従う
- アプリバー改修・類題ボタンはしない
- 変更は mos_xaml_app 配下（＋ MOSTest の Excel ファイル）のみ

原因 → 修正 → ビルド → 再確認手順 を報告してください。
```

---

## 7. 作業後チェックリスト

### 先行体験（§1）後

- [ ] `ExcelChecker*_PV1.cs` が新規追加されている
- [ ] タスク1だけ類題化、他は委譲
- [ ] `MOS演習問題文一覧.json` の **taskId:1 だけ** 類題文言になっている
- [ ] アプリバーに類題の問題文が表示される
- [ ] 教材 Checker は上書きしていない
- [ ] `config.json` の `library` が `_PV1` になっている
- [ ] Debug ビルドが通る
- [ ] VS でタスク1の **〇/✖** を確認した

### 本番（§2）後

- [ ] `Tab1\Initial\project{N}.xlsx` が類題内容になっている
- [ ] `ExcelChecker*_PV*.cs` が **新規追加** され、クラス名が `ExcelChecker` で始まる
- [ ] 教材 Checker ファイルは上書きしていない（参照のみ）
- [ ] Terminal の copy/cp で Checker を複製していない
- [ ] `config.json` の `library` が土台ガイド §3 の「類題1 Checker」と一致している
- [ ] `MOS演習問題文一覧.json` が類題文言になっている
- [ ] Debug ビルドが通る
- [ ] AppBar / MainViewModel に類題ボタン・モード切替を **追加していない**
- [ ] ランチャー config を触っていない

---

## 8. 依頼のコツ

| やり方 | おすすめ度 | 説明 |
|--------|------------|------|
| §1 先行体験 → §2 本番の順 | ◎◎ | いちばん失敗しにくい |
| §2 本番だけ一気に | ◎ | 経験者向け |
| Step 別プロンプト（§4）で分割 | ◎◎ | 本番を細かく進めたいとき |
| 「全部やって」だけ | △ | 範囲が広がりやすい |

---

## 改訂履歴

| 日付 | 内容 |
|------|------|
| 2026-06-12 | 初版 |
| 2026-06-12 | 類題ボタン・アプリ改修のプロンプトを削除。作成・VS確認動線のみに限定 |
| 2026-06-12 | 例をプロジェクト1基準に修正。プロジェクト2以降は追加オプション 2-D に整理 |
| 2026-06-12 | §3 全プロジェクト対応表への参照を追加。2-E（プロジェクト7）を追加 |
| 2026-06-12 | Checker は Terminal 複製ではなく新規作成を依頼する方針を明記 |
| 2026-06-12 | §1 先行体験プロンプト追加。§2 本番に分離。手順の順序を明確化 |
| 2026-06-12 | 先行体験に問題文（taskId:1 のみ）作成を含める |
