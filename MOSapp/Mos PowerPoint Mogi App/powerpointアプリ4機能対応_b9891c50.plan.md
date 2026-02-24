---
name: PowerPointアプリ4機能対応
overview: レビューページから問題クリック時にPowerPointを起動する、リセット後の読み取り専用を防ぐ、プロジェクト一覧に「すべてをリセットする」を追加、次のプロジェクト押下時に日時付きバックアップ保存を行う4点を実装する計画です。
todos:
  - id: todo-1771341004714-h9y40243b
    content: "プロジェクト一覧画面に「バックアップ確認」ボタンを追加。機能は保存先ディレクトリ: {basePath}/Tab{_groupId}/backup/{MMDD_HHmm}/ とリンク"
    status: pending
isProject: false
---

# PowerPointアプリ 4機能対応プラン

## 1. レビューページで問題クリック時にPowerPointを起動する

**現象**: レビューページで問題文（例: projectId=9, taskId=2）をクリックするとアプリバーにはその問題が表示されるが、PowerPointが起動・切り替わらない。

**原因**: [Views/UiTestAppBarWindow.xaml.cs](MOSapp/Mos PowerPoint Mogi App/Views/UiTestAppBarWindow.xaml.cs) の `NavigateToTask` で、プロジェクト変更時に `_currentProjectId` の更新と `LoadCurrentProjectTasks()` / `UpdateTaskDisplay()` のみ行っており、`**OpenProjectDocument(projectId, _groupId)` を呼んでいない**ため。

**対応**:

- `NavigateToTask` 内で、`projectId != _currentProjectId` のとき、`OpenProjectDocument(projectId, _groupId)` を呼ぶ（既存の `MoveToNextProject` と同様に、該当プロジェクトの pptx を開く）。
- 他プロジェクトのプレゼンを開く前に、既存の `OpenProjectDocument` と同様に「開いているプレゼンを閉じてから開く」処理がそのまま適用されるため、追加で閉じる処理は不要（`OpenProjectDocument` 内で既に実施済み）。

---

## 2. プロジェクトをリセットすると読み取り専用にならないようにする

**原因**: [Views/UiTestAppBarWindow.xaml.cs](MOSapp/Mos PowerPoint Mogi App/Views/UiTestAppBarWindow.xaml.cs) の `ResetProject` で、テンプレートを読み取り専用にしたうえで `File.Copy(templatePath, projectFilePath, overwrite: true)` している。**コピー先（projectFilePath）がテンプレートの読み取り専用属性を引き継ぎ、リセット後のプロジェクトファイルが読み取り専用になる**可能性がある。現状はコピー「前」にコピー先の読み取り専用解除のみで、コピー「後」の明示的な解除がない。

**対応**:

- `File.Copy` の**直後**に、`projectFilePath` の `FileInfo.IsReadOnly = false` を設定する処理を追加する。
- 同様に、Initial フォルダへコピーした `initialFilePath` についても、コピー成功後に読み取り専用を解除する（既に存在時のみ解除している箇所の近くで、コピー後にも確実に解除する）。

---

## 3. プロジェクト一覧画面に「すべてをリセットする」を追加する

**対象画面**: [MainWindow.xaml](MOSapp/MOSapp/Mos PowerPoint Mogi App/MainWindow.xaml) の TabControl（演習 / 応用編）で表示しているプロジェクト一覧。「すべてをリセットする」はこのメイン画面に配置する。

**機能**: 全グループ（Tab1=演習, Tab3=応用編）の全プロジェクト（Project1～11）について、`ResetProject(groupId, projectId)` と同等の処理を行う（テンプレートからプロジェクトファイルを上書きし、Initial にもコピー。PowerPoint は起動していない状態で実行する想定）。

**実装方針**:

- **リセットロジックの共通化**: `ResetProject` の本体は [UiTestAppBarWindow.xaml.cs](MOSapp/Mos PowerPoint Mogi App/Views/UiTestAppBarWindow.xaml.cs) にあり、パス解決・テンプレート検索・コピー・Initial コピー・読み取り専用解除が含まれる。これを **static メソッド**（例: `PowerPointProjectResetHelper.ResetProject(int groupId, int projectId)`）に切り出し、UiTestAppBarWindow の `ResetProject` と「すべてをリセット」の両方から呼ぶ形にする。
  - 新規クラス例: `PowerPointProjectResetHelper.cs` をプロジェクトに追加し、`App.config` の `PowerPointDataPath` を使う同一ロジックを static で実装。UiTestAppBarWindow の `ResetProject` はその static を呼ぶだけにする。
- **MainWindow 側**: MainViewModel に `ResetAllProjectsCommand` を追加し、実行時に「すべてのプロジェクトをリセットしてよいか」確認ダイアログを表示したうえで、groupId ∈ {1, 3} と projectId ∈ {1..11} の組み合わせで `PowerPointProjectResetHelper.ResetProject(groupId, projectId)` を順次呼ぶ。成功/失敗メッセージは最後にまとめて表示する。
- **UI**: MainWindow.xaml のプロジェクト一覧がある Tab の上部または下部に「すべてをリセットする」ボタンを追加し、`ResetAllProjectsCommand` にバインドする。

---

## 4. 「次のプロジェクト」クリック時に日付・時間でバックアップ保存する

**要件**: 「次のプロジェクト」ボタン押下時に、**現在開いている PowerPoint を日付・時間付きで保存し、backup フォルダを作成してそこに自動保存する**。フォーマット例: `0218_1035`（MMDD_HHmm）。

**実装方針**:

- [Views/UiTestAppBarWindow.xaml.cs](MOSapp/Mos PowerPoint Mogi App/Views/UiTestAppBarWindow.xaml.cs) の `MoveToNextProject` の**先頭**で、現在のプロジェクト（`_currentProjectId`）のプレゼンテーションが開いていれば、以下を行う:
  1. 保存先ディレクトリ: `{basePath}/Tab{_groupId}/backup/{MMDD_HHmm}/` を作成（例: `C:\MOSTest\PowerPoint365\Tab1\backup\0218_1035\`）。
  2. 現在の PowerPoint のプレゼンテーションを `SaveAs` で上記フォルダに保存（ファイル名は現在のプロジェクトに合わせて `Project{_currentProjectId}.pptx` など）。
  3. 保存後、既存の「開いているプレゼンを閉じる → 次のプロジェクトを開く」の流れ（`OpenProjectDocument(_currentProjectId, _groupId)`）に進む。
- 日時フォーマット: `DateTime.Now` で `"MMdd_HHmm"` を使用（例: 2月18日 10:35 → `0218_1035`）。
- PowerPoint が起動していない、またはプレゼンテーションが開いていない場合は、バックアップをスキップして従来どおり次のプロジェクトを開くだけにする。

---

## 変更ファイル一覧（想定）


| 対象                                                                                            | 変更内容                                                                                                                                                                                                                               |
| --------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [UiTestAppBarWindow.xaml.cs](MOSapp/Mos PowerPoint Mogi App/Views/UiTestAppBarWindow.xaml.cs) | ① NavigateToTask でプロジェクト変更時に OpenProjectDocument を呼ぶ ② ResetProject でコピー後に projectFilePath / initialFilePath の読み取り専用を明示的に解除 ③ ResetProject 本体を PowerPointProjectResetHelper に委譲 ④ MoveToNextProject の先頭で backup/MMdd_HHmm に SaveAs |
| 新規: PowerPointProjectResetHelper.cs                                                           | ⑤ ResetProject の static 版（パス・テンプレート検索・コピー・読み取り専用解除）を実装                                                                                                                                                                             |
| [MainWindow.xaml](MOSapp/MOSapp/Mos PowerPoint Mogi App/MainWindow.xaml)                      | ⑥ 「すべてをリセットする」ボタン追加                                                                                                                                                                                                                |
| [MainViewModel.cs](MOSapp/MOSapp/Mos PowerPoint Mogi App/MainViewModel.cs)                    | ⑦ ResetAllProjectsCommand の追加と、全 groupId/projectId に対する ResetProject 呼び出し                                                                                                                                                          |


---

## 補足

- バックアップは「次のプロジェクト」押下時のみ実行し、タイマーで自動遷移した場合（5分経過で次のプロジェクト）も同様にバックアップを取るかは仕様次第。必要なら `MoveToNextProjectWithMessage` → `MoveToNextProject` の流れで同じ `MoveToNextProject` を使うため、そのままバックアップも実行される。
- リセット時の「読み取り専用」は、テンプレート自体は読み取り専用のままにして、**コピー先のプロジェクトファイルと Initial コピー先だけ確実に読み取り専用解除**する方針で対応する。

