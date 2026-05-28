using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Newtonsoft.Json;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WordDoc = Microsoft.Office.Interop.Word.Document;

using Libraries.Group1;

namespace Libraries
{
    /// <summary>
    /// 試験終了時の全プロジェクト一括採点（WordChecker DLL 利用）。
    /// </summary>
    public static class WordBatchScoring
    {
        private class BatchProjectData
        {
            [JsonProperty("projects")]
            public List<BatchProjectInfo> Projects { get; set; }
        }

        private class BatchProjectInfo
        {
            [JsonProperty("projectId")]
            public int ProjectId { get; set; }
            [JsonProperty("tasks")]
            public List<BatchTaskInfo> Tasks { get; set; }
        }

        private class BatchTaskInfo
        {
            [JsonProperty("taskId")]
            public int TaskId { get; set; }
        }

        /// <summary>
        /// 問題文 JSON を読み込み、全プロジェクトを採点して ScoreResultStore に保存する。
        /// </summary>
        public static void ScoreAllProjects(int groupId)
        {
            var projectData = LoadProjectData();
            if (projectData?.Projects == null || projectData.Projects.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[WordBatchScoring] No project data loaded");
                return;
            }

            ScoreResultStore.ClearGroup(groupId);

            WordApp wordApp = null;
            try
            {
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    wordApp = new WordApp();
                    Thread.Sleep(500);
                    try { wordApp.Visible = true; } catch { }
                }

                if (wordApp != null)
                {
                    try { wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone; } catch { }
                    try { wordApp.Visible = true; } catch { }
                    WordWindowLayoutHelper.PositionWordForBatchScoring(wordApp);
                }

                foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
                {
                    int taskCount = project.Tasks?.Count ?? 0;
                    if (taskCount == 0)
                        continue;

                    System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] Scoring project {project.ProjectId} ({taskCount} tasks)");

                    if (!OpenProjectDocument(wordApp, project.ProjectId, groupId))
                    {
                        for (int t = 1; t <= taskCount; t++)
                            ScoreResultStore.RecordResult(groupId, project.ProjectId, t, false);
                        continue;
                    }

                    Thread.Sleep(800);

                    var results = ScoreProject(groupId, project.ProjectId, taskCount);
                    for (int i = 0; i < taskCount; i++)
                    {
                        bool passed = i < results.Count && results[i];
                        ScoreResultStore.RecordResult(groupId, project.ProjectId, i + 1, passed);
                    }
                }
            }
            finally
            {
                if (wordApp != null)
                {
                    try { Marshal.ReleaseComObject(wordApp); } catch { }
                }
            }

            System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] Done scoring group {groupId}");
        }

        /// <summary>
        /// 1 問だけ再採点する。成功時は結果を返し、失敗時は null。
        /// </summary>
        public static bool? ScoreSingleTask(int groupId, int projectId, int taskId)
        {
            WordApp wordApp = null;
            try
            {
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    wordApp = new WordApp();
                    Thread.Sleep(500);
                    try { wordApp.Visible = true; } catch { }
                }

                if (wordApp == null)
                    return null;

                try { wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone; } catch { }

                SaveAllOpenDocuments(wordApp);

                // 復習中の編集を捨てない: 既に開いている場合は保存してそのまま採点する
                if (!TryActivateOpenProjectDocument(wordApp, projectId, groupId)
                    && !OpenProjectDocument(wordApp, projectId, groupId))
                {
                    System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] ScoreSingleTask: could not open project {projectId}");
                    return null;
                }

                Thread.Sleep(300);
                int attemptNo = WordTaskAttemptRegistry.GetAttempt(projectId, taskId);
                if (!WordGradingGate.TryPass(groupId, projectId, taskId, attemptNo, out _))
                    return false;
                bool? result = InvokeCheckTask(groupId, projectId, taskId);
                if (result.HasValue)
                    ScoreResultStore.RecordResult(groupId, projectId, taskId, result.Value);
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] ScoreSingleTask error: {ex.Message}");
                return null;
            }
            finally
            {
                if (wordApp != null)
                {
                    try { Marshal.ReleaseComObject(wordApp); } catch { }
                }
            }
        }

        private static bool? InvokeCheckTask(int groupId, int projectId, int taskNum)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string dllPath = Path.Combine(baseDir, "Dlls", $"WordChecker{groupId}_{projectId}.dll");
            if (!File.Exists(dllPath))
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] DLL not found: {dllPath}");
                return null;
            }

            try
            {
                Assembly assembly = Assembly.LoadFrom(dllPath);
                string className = $"Libraries.Group{groupId}.WordChecker{groupId}_{projectId}";
                Type checkerType = assembly.GetType(className);
                if (checkerType == null)
                    return null;

                object checkerInstance = Activator.CreateInstance(checkerType);
                string methodName = $"CheckTask_{groupId}_{projectId}_{taskNum:D2}";
                MethodInfo method = checkerType.GetMethod(methodName);
                if (method == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] P{projectId} T{taskNum}: method not found ({methodName})");
                    return null;
                }

                bool taskResult = (bool)method.Invoke(checkerInstance, null);
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] ScoreSingleTask P{projectId} T{taskNum}: {(taskResult ? "pass" : "fail")}");
                return taskResult;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] InvokeCheckTask P{projectId} T{taskNum} error: {ex.Message}");
                return null;
            }
        }

        private static BatchProjectData LoadProjectData()
        {
            try
            {
                string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", "MOS模擬アプリ問題文一覧_Word.json");
                if (!File.Exists(jsonPath))
                    jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MOS模擬アプリ問題文一覧_Word.json");
                if (!File.Exists(jsonPath))
                    return null;

                string jsonContent = File.ReadAllText(jsonPath, Encoding.UTF8);
                return JsonConvert.DeserializeObject<BatchProjectData>(jsonContent);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] LoadProjectData error: {ex.Message}");
                return null;
            }
        }

        private static List<bool> ScoreProject(int groupId, int projectId, int taskCount)
        {
            var results = new List<bool>();
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string dllPath = Path.Combine(baseDir, "Dlls", $"WordChecker{groupId}_{projectId}.dll");

            if (!File.Exists(dllPath))
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] DLL not found: {dllPath}");
                for (int i = 0; i < taskCount; i++)
                    results.Add(false);
                return results;
            }

            try
            {
                Assembly assembly = Assembly.LoadFrom(dllPath);
                string className = $"Libraries.Group{groupId}.WordChecker{groupId}_{projectId}";
                Type checkerType = assembly.GetType(className);

                if (checkerType == null)
                {
                    for (int i = 0; i < taskCount; i++)
                        results.Add(false);
                    return results;
                }

                object checkerInstance = Activator.CreateInstance(checkerType);

                for (int taskNum = 1; taskNum <= taskCount; taskNum++)
                {
                    string methodName = $"CheckTask_{groupId}_{projectId}_{taskNum:D2}";
                    MethodInfo method = checkerType.GetMethod(methodName);

                    if (method == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] P{projectId} T{taskNum}: method not found ({methodName})");
                        results.Add(false);
                        continue;
                    }

                    try
                    {
                        int attemptNo = 0;
                        if (!WordGradingGate.TryPass(groupId, projectId, taskNum, attemptNo, out _))
                        {
                            results.Add(false);
                            System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] P{projectId} T{taskNum}: fail (grading gate)");
                            continue;
                        }
                        bool taskResult = (bool)method.Invoke(checkerInstance, null);
                        results.Add(taskResult);
                        System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] P{projectId} T{taskNum}: {(taskResult ? "pass" : "fail")}");
                    }
                    catch (Exception exTask)
                    {
                        System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] P{projectId} T{taskNum} error: {exTask.Message}");
                        results.Add(false);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] ScoreProject error: {ex.Message}");
                while (results.Count < taskCount)
                    results.Add(false);
            }

            return results;
        }

        /// <summary>
        /// 開いている全 Word 文書を保存する（再採点前にユーザーの編集を残す）。
        /// </summary>
        private static void SaveAllOpenDocuments(WordApp wordApp)
        {
            if (wordApp == null) return;
            try
            {
                for (int i = wordApp.Documents.Count; i >= 1; i--)
                {
                    WordDoc doc = null;
                    try
                    {
                        doc = wordApp.Documents[i];
                        if (!doc.Saved)
                            doc.Save();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] SaveAllOpenDocuments: {ex.Message}");
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try { Marshal.ReleaseComObject(doc); } catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] SaveAllOpenDocuments error: {ex.Message}");
            }
        }

        /// <summary>
        /// 対象プロジェクトのファイルが既に開いていれば保存して前面化する（閉じ直ししない）。
        /// </summary>
        private static bool TryActivateOpenProjectDocument(WordApp wordApp, int projectId, int groupId)
        {
            if (wordApp == null)
                return false;

            string filePath = ResolveProjectFilePath(projectId, groupId);
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return false;

            try
            {
                string pathLower = Path.GetFullPath(filePath).ToLowerInvariant();
                for (int i = wordApp.Documents.Count; i >= 1; i--)
                {
                    WordDoc doc = null;
                    try
                    {
                        doc = wordApp.Documents[i];
                        string fullName = doc.FullName?.ToLowerInvariant() ?? "";
                        string docFullPath = fullName;
                        try { docFullPath = Path.GetFullPath(fullName).ToLowerInvariant(); } catch { }
                        if (fullName != pathLower && docFullPath != pathLower)
                            continue;

                        if (!doc.Saved)
                            doc.Save();
                        doc.Activate();
                        System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] Reusing open document for P{projectId} (rescore)");
                        return true;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] TryActivateOpenProjectDocument: {ex.Message}");
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try { Marshal.ReleaseComObject(doc); } catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] TryActivateOpenProjectDocument error: {ex.Message}");
            }

            return false;
        }

        private static bool OpenProjectDocument(WordApp wordApp, int projectId, int groupId)
        {
            if (wordApp == null)
                return false;

            string filePath = ResolveProjectFilePath(projectId, groupId);
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] File not found for project {projectId}");
                return false;
            }

            try
            {
                string pathLower = Path.GetFullPath(filePath).ToLowerInvariant();

                for (int i = wordApp.Documents.Count; i >= 1; i--)
                {
                    WordDoc doc = null;
                    try
                    {
                        doc = wordApp.Documents[i];
                        string fullName = doc.FullName?.ToLowerInvariant() ?? "";
                        string docFullPath = fullName;
                        try { docFullPath = Path.GetFullPath(fullName).ToLowerInvariant(); } catch { }
                        if (fullName == pathLower || docFullPath == pathLower)
                        {
                            // 同一ファイルが既に開いている場合は、変更を捨てずに保存して再利用する。
                            // Project7 は SaveAs 系タスクがあるため、ここでの Save() により
                            // 「名前を付けて保存」UI が出るケースを避ける。
                            if (projectId != 7)
                            {
                                try
                                {
                                    if (!doc.Saved)
                                        doc.Save();
                                }
                                catch { }
                            }
                            try
                            {
                                doc.Activate();
                            }
                            catch { }
                            Thread.Sleep(200);
                            WordWindowLayoutHelper.PositionWordForBatchScoring(wordApp);
                            Thread.Sleep(100);
                            return true;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (doc != null) Marshal.ReleaseComObject(doc);
                    }
                }

                WordDoc opened = wordApp.Documents.Open(filePath, ReadOnly: false, Visible: true);
                try
                {
                    opened.Activate();
                }
                catch { }
                finally
                {
                    if (opened != null) Marshal.ReleaseComObject(opened);
                }
                Thread.Sleep(300);
                WordWindowLayoutHelper.PositionWordForBatchScoring(wordApp);
                Thread.Sleep(200);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] OpenProjectDocument error: {ex.Message}");
                return false;
            }
        }

        private static string ResolveProjectFilePath(int projectId, int groupId)
        {
            string basePath = @"C:\MOSTest\Word365";
            string workingFolder = Path.Combine(basePath, $"Tab{groupId}");
            string initialFolder = Path.Combine(basePath, $"Tab{groupId}", "Initial");
            string initialInitialFolder = Path.Combine(basePath, $"Tab{groupId}", "Initial", "Initial");
            string[] possibleNames = (groupId == 1 && projectId == 7)
                ? new[] { $"Project{projectId}.doc", $"project{projectId}.doc" }
                : new[] { $"Project{projectId}.docx", $"Project{projectId}.doc", $"project{projectId}.docx", $"project{projectId}.doc" };
            string workingFileName = (groupId == 1 && projectId == 7) ? "Project7.doc" : $"Project{projectId}.docx";
            string workingFilePath = Path.Combine(workingFolder, workingFileName);

            foreach (var fileName in possibleNames)
            {
                string fullPath = Path.Combine(workingFolder, fileName);
                if (File.Exists(fullPath))
                    return fullPath;
            }

            string sourcePath = null;
            foreach (var fileName in possibleNames)
            {
                string fullPath = Path.Combine(initialFolder, fileName);
                if (File.Exists(fullPath)) { sourcePath = fullPath; break; }
            }
            if (string.IsNullOrEmpty(sourcePath) && Directory.Exists(initialInitialFolder))
            {
                foreach (var fileName in possibleNames)
                {
                    string fullPath = Path.Combine(initialInitialFolder, fileName);
                    if (File.Exists(fullPath)) { sourcePath = fullPath; break; }
                }
            }

            if (!string.IsNullOrEmpty(sourcePath))
            {
                try
                {
                    if (!Directory.Exists(workingFolder))
                        Directory.CreateDirectory(workingFolder);
                    File.Copy(sourcePath, workingFilePath, overwrite: false);
                    return workingFilePath;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] Copy from Initial failed: {ex.Message}");
                }
            }

            return null;
        }
    }

    // -------------------------------------------------------------------------
    // 破壊的操作検知（採点前ゲート WordGradingGate）
    //
    // 【記録の種類】
    // ・スナップショット … VSTO がタスク開始時に Project{N} の COM 状態を mos_word_snapshot.txt に保存。
    //   タスク切替時・採点時に差分 → mos_word_destructive_errors.log（例: SectionsCount, BodyTextLength）。
    //   リボン未接続の編集（余白・段組み・書式など）の補完に使う。
    // ・[Op] … VSTO Ribbon / ポーリングが mos_word_log.txt に [Task P-T-A] [Op] Type … で追記。
    //   例: Cut, Paste, InsertTable, FileSaveAs, HeaderFooterEdit, RibbonCommand（idMso 汎用）。
    // ・Executed … 正解証跡（破壊検知とは別軸）。同じ mos_word_log.txt。[ProjectN] … Executed。
    // ・TaskStart … 試験アプリがタスク表示時に記録。Op 区間の境界に使用。
    //
    // 【採点ゲート順】①既存 destructive ログ → ②許可外 [Op] → ③スナップショット差分 → ④WordChecker
    // -------------------------------------------------------------------------

    /// <summary>
    /// スナップショット差分で「変更してよい」項目。タスク表示時に current_task の第3列へ書き出し、VSTO 比較でも参照。
    /// </summary>
    [Flags]
    public enum WordValidationExemptFlags
    {
        None = 0,
        /// <summary>［レイアウト］区切り・次のページから開始 等（セクション数）</summary>
        SectionsCount = 1,
        /// <summary>切り取り・貼り付け・置換・入力など（本文の Character 数）</summary>
        BodyTextLength = 2,
        /// <summary>［挿入］画像（行内）等（InlineShapes 数）</summary>
        InlineShapesCount = 4,
        /// <summary>図形・テキストボックス等（Shapes 数）</summary>
        FloatingShapesCount = 8,
        /// <summary>［挿入］表・文字列を表にする 等（Tables 数）</summary>
        TablesCount = 16,
        /// <summary>［校閲］新しいコメント・返信・解決・削除 等（Comments 数）</summary>
        CommentsCount = 32,
        /// <summary>［挿入］ヘッダー／フッター、ドキュメント検査での削除 等（先頭セクション primary ヘッダー指紋）</summary>
        HeaderFooterFingerprint = 64,
        /// <summary>［デザイン］ページ罫線（PageBorder の有無は将来拡張。現状は差分キーのみ）</summary>
        PageBorder = 128,
        /// <summary>［デザイン］透かし</summary>
        Watermark = 256,
        /// <summary>［ホーム］編集記号の表示/非表示（非表示記号の表示状態）</summary>
        ShowAllState = 512,
        /// <summary>［ファイル］情報の「変換」（互換モード）</summary>
        CompatibilityMode = 1024,
        /// <summary>7-4/7-5: 朗読会.txt・朗読会.docm 等の同フォルダ派生ファイルの出現</summary>
        SiblingFiles = 2048,
        /// <summary>7-4/7-5: アクティブ文書が Project7.doc 以外（朗読会.txt / .docm）への切替</summary>
        ActiveDocumentSwitch = 4096
    }

    /// <summary>
    /// タスク別のスナップショット免除と [Op] 許可種別。未設定タスクは免除なし・許可 Op は RibbonCommand のみ。
    /// </summary>
    public static class WordTaskValidationConfig
    {
        /// <summary>スナップショット比較で無視する変更（正解操作で必ず変わる項目）。</summary>
        public static WordValidationExemptFlags GetExemptFlags(int projectId, int taskId)
        {
            WordValidationExemptFlags flags = WordValidationExemptFlags.None;

            // P1
            if (projectId == 1 && taskId == 1)
                // 1-1: ［ホーム］編集記号の表示/非表示（表示→非表示→表示）
                flags |= WordValidationExemptFlags.ShowAllState;

            // P2
            if (projectId == 2 && taskId == 1)
                // 2-1: ［ホーム］切り取り・貼り付け（本文移動）
                flags |= WordValidationExemptFlags.BodyTextLength;

            // P3
            if (projectId == 3 && taskId == 2)
            {
                // 3-2: ［レイアウト］区切り→次のページから開始（セクション区切り記号で BodyTextLength +1 し得る）
                flags |= WordValidationExemptFlags.SectionsCount | WordValidationExemptFlags.BodyTextLength;
            }
            if (projectId == 3 && taskId == 5)
                // 3-5: ［レイアウト］段区切り（列区切り Chr(14) で BodyTextLength +1 し得る）
                flags |= WordValidationExemptFlags.BodyTextLength;
            // 3-1, 3-3, 3-4, 3-6: 上記以外は SectionsCount/BodyTextLength 免除なし（誤検知時は個別に再付与）

            // P4
            if (projectId == 4)
            {
                if (taskId >= 1 && taskId <= 3)
                    // 4-1〜4-3: ［校閲］コメント挿入・返信・解決・削除
                    flags |= WordValidationExemptFlags.CommentsCount;
                if (taskId == 5)
                    // 4-5: ［デザイン］透かし（下書き1）— タスク区間内の Watermark 変化のみ免除
                    flags |= WordValidationExemptFlags.Watermark;
                if (taskId == 6)
                    // 4-6: ページ罫線 — タスク区間内の PageBorder 変化のみ免除
                    flags |= WordValidationExemptFlags.PageBorder;
                if (taskId == 7)
                    // 4-7: ドキュメント検査でヘッダー・フッター・透かし削除
                    flags |= WordValidationExemptFlags.HeaderFooterFingerprint | WordValidationExemptFlags.Watermark;
            }

            // P5
            if (projectId == 5)
            {
                // 5-1〜5-8: ［挿入］画像、図の効果、代替テキスト、背景の削除 等
                flags |= WordValidationExemptFlags.InlineShapesCount | WordValidationExemptFlags.FloatingShapesCount;
                if (taskId == 1 || taskId == 2)
                    // 5-1/5-2: 行内挿入・四角形化で doc.Content の文字数が ±1 し得る（正答の副作用）
                    flags |= WordValidationExemptFlags.BodyTextLength;
            }

            // P6
            if (projectId == 6)
            {
                // 6-1/2/4〜7: 表作成・分割・表挿入で TablesCount が変わり得る。6-3 はセル分割のみで表数は不変想定
                if (taskId != 3)
                    flags |= WordValidationExemptFlags.TablesCount;
                // 6-1/2/3/7: 文字列を表にする・セル入力で本文長が変わり得る。6-4/5/6 は書式のみで BodyTextLength 免除なし
                if (taskId != 4 && taskId != 5 && taskId != 6)
                    flags |= WordValidationExemptFlags.BodyTextLength;
            }

            // P7（タスク 1〜5）
            if (projectId == 7)
            {
                // 7-1〜7-5: 7-1［ファイル］情報の「変換」による互換モード変化を後続でも許容
                flags |= WordValidationExemptFlags.CompatibilityMode;
                if (taskId == 3)
                    // 7-3: ［挿入］ヘッダー→インテグラル（ヘッダー内容の変化は正解）
                    flags |= WordValidationExemptFlags.HeaderFooterFingerprint;
                if (taskId == 4 || taskId == 5)
                {
                    // 7-4: 名前を付けて保存→書式なし(.txt)／7-5: マクロ有効(.docm)・読み取りパスワード
                    flags |= WordValidationExemptFlags.ActiveDocumentSwitch | WordValidationExemptFlags.SiblingFiles;
                }
            }

            // P8
            if (projectId == 8)
            {
                if (taskId == 2 || taskId == 4)
                    // 8-2: 脚注入力／8-4: § 特殊文字挿入で doc.Content 長が +1 し得る
                    flags |= WordValidationExemptFlags.BodyTextLength;
                if (taskId == 5)
                    // 8-5: SmartArt 挿入で行内図形数・本文長が変わり得る（8-6 は色のみで免除なし）
                    flags |= WordValidationExemptFlags.BodyTextLength | WordValidationExemptFlags.InlineShapesCount;
            }

            // P9
            if (projectId == 9)
            {
                if (taskId == 2)
                    // 9-2: 表の解除（コンマ区切り）で表数・本文長が変わる
                    flags |= WordValidationExemptFlags.TablesCount | WordValidationExemptFlags.BodyTextLength;
                if (taskId == 5)
                    // 9-5: すべての変更履歴を反映で本文長が変わり得る
                    flags |= WordValidationExemptFlags.BodyTextLength;
            }

            // P10
            if (projectId == 10 && taskId == 1)
                // 10-1: 「ウイルス」→「コンピュータウイルス」の一括置換で doc.Content 長が増える
                flags |= WordValidationExemptFlags.BodyTextLength;

            return flags;
        }

        /// <summary>
        /// mos_word_log の [Op] で許可する操作種別。それ以外（Cut 等）は採点ゲート②で ✖。
        /// 全タスク共通で RibbonCommand（リボン idMso の汎用記録）を許可。
        /// </summary>
        public static HashSet<string> GetAllowedOperationTypes(int projectId, int taskId)
        {
            // 記録元: VSTO Ribbon.OnActionCallback → Logger.LogOperation（未接続はスナップショットのみ）
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "RibbonCommand" };

            // P2
            if (projectId == 2 && taskId == 1)
            {
                // 2-1: ［ホーム］切り取り・貼り付け — [Op] 明示記録
                allowed.Add("Cut");
                allowed.Add("Paste");
            }
            if (projectId == 2 && taskId == 4)
                // 2-4: 問題文の指定文言クリックでコピー → 図形へ［貼り付け］（直接入力より Paste 想定）
                allowed.Add("Paste");

            // P3
            if (projectId == 3 && taskId == 2)
                // 3-2: ［レイアウト］区切り — [Op] InsertSectionBreak
                allowed.Add("InsertSectionBreak");

            // P4
            if (projectId == 4 && taskId >= 1 && taskId <= 3)
                // 4-1〜4-3: ［校閲］新しいコメント 等 — [Op] ReviewNewComment
                allowed.Add("ReviewNewComment");
            if (projectId == 4 && (taskId == 1 || taskId == 2))
            {
                // 4-1/4-2: 問題文の指定文言クリックでコピー → コメント欄へ［貼り付け］（2-4 と同様）
                allowed.Add("Paste");
            }
            if (projectId == 4 && taskId == 5)
            {
                // 4-5: ［デザイン］透かし（下書き1）— [Op] Watermark（VSTO ポーリング。他タスクでは許可外）
                allowed.Add("Watermark");
            }
            if (projectId == 4 && taskId == 6)
                // 4-6: ［デザイン］ページ罫線 — [Op] PageBorders（Ribbon / ポーリング。他タスクでは許可外）
                allowed.Add("PageBorders");

            // P5
            if (projectId == 5)
                // P5 全般: ［挿入］このデバイスから画像 — [Op] InsertPicture（主にスナップショットで検証）
                allowed.Add("InsertPicture");

            // P6
            if (projectId == 6)
                // P6 全般: ［挿入］表 — [Op] InsertTable
                allowed.Add("InsertTable");
            if (projectId == 6 && taskId == 7)
                // 6-7: 問題文の指定文言クリックでコピー → セルへ［貼り付け］（2-4/4-1/4-2 と同様）
                allowed.Add("Paste");

            // P7
            if (projectId == 7 && taskId == 3)
                // 7-3: ［挿入］ヘッダー編集 — [Op] HeaderFooterEdit（正解は Executed IntegralHeader も併用）
                allowed.Add("HeaderFooterEdit");
            if (projectId == 7 && (taskId == 4 || taskId == 5))
                // 7-4/7-5: ［ファイル］名前を付けて保存 — [Op] およびポーリング FileSaveAsTxt / FileSaveAsDocm（Executed 側）
                allowed.Add("FileSaveAs");

            // P8
            if (projectId == 8 && (taskId == 2 || taskId == 5))
            {
                // 8-2: 脚注／8-5: SmartArt テキスト — 問題文クリックでコピー → ［貼り付け］（2-4/4-1/4-2 と同様）
                allowed.Add("Paste");
            }

            return allowed;
        }
    }

    /// <summary>
    /// タスク開始時スナップショット（mos_word_snapshot.txt）と採点時の COM 再取得の差分。
    /// 取得: VSTO WordDestructiveMonitor.TakeSnapshot／採点: 本クラス CompareAndGetErrors。
    /// </summary>
    public static class WordSnapshotChecker
    {
        public class SnapshotData
        {
            public int ProjectId { get; set; }
            public int TaskId { get; set; }
            public int AttemptNo { get; set; }
            public string FullName { get; set; }
            public int Sections { get; set; }
            public int BodyTextLength { get; set; }
            public int InlineShapes { get; set; }
            public int FloatingShapes { get; set; }
            public int Tables { get; set; }
            public int Comments { get; set; }
            public string HeaderPrimaryFp { get; set; }
            public int CompatibilityMode { get; set; } = -1;
            public string WatermarkFingerprint { get; set; } = "None";
            public string PageBorderFingerprint { get; set; } = "";
        }

        public static List<string> CompareAndGetErrors(int groupId, int projectId, int taskId, int attemptNo, WordValidationExemptFlags exemptFlags)
        {
            var errors = new List<string>();
            var snapshot = LoadSnapshot();
            if (snapshot == null)
                return errors;
            if (snapshot.ProjectId != projectId || snapshot.TaskId != taskId || snapshot.AttemptNo != attemptNo)
                return errors;

            SnapshotData current = CaptureProjectDocument(projectId, groupId);
            if (current == null)
                return errors;

            // 以下はいずれもスナップショット専用（[Op] 未記録の操作を補完）。文言は mos_word_destructive_errors.log にそのまま出る。
            if (!exemptFlags.HasFlag(WordValidationExemptFlags.SectionsCount) && current.Sections != snapshot.Sections)
                errors.Add($"SectionsCount changed {snapshot.Sections}->{current.Sections}"); // 例: ［レイアウト］区切り
            if (!exemptFlags.HasFlag(WordValidationExemptFlags.BodyTextLength) && current.BodyTextLength != snapshot.BodyTextLength)
                errors.Add($"BodyTextLength changed {snapshot.BodyTextLength}->{current.BodyTextLength}"); // 例: 入力・切り取り・置換
            if (!exemptFlags.HasFlag(WordValidationExemptFlags.InlineShapesCount) && current.InlineShapes != snapshot.InlineShapes)
                errors.Add($"InlineShapesCount changed {snapshot.InlineShapes}->{current.InlineShapes}"); // 例: ［挿入］画像（行内）
            if (!exemptFlags.HasFlag(WordValidationExemptFlags.FloatingShapesCount) && current.FloatingShapes != snapshot.FloatingShapes)
                errors.Add($"FloatingShapesCount changed {snapshot.FloatingShapes}->{current.FloatingShapes}"); // 例: 図形・テキストボックス
            if (!exemptFlags.HasFlag(WordValidationExemptFlags.TablesCount) && current.Tables != snapshot.Tables)
                errors.Add($"TablesCount changed {snapshot.Tables}->{current.Tables}"); // 例: ［挿入］表
            if (!exemptFlags.HasFlag(WordValidationExemptFlags.CommentsCount) && current.Comments != snapshot.Comments)
                errors.Add($"CommentsCount changed {snapshot.Comments}->{current.Comments}"); // 例: ［校閲］コメント
            if (!exemptFlags.HasFlag(WordValidationExemptFlags.HeaderFooterFingerprint)
                && !string.Equals(snapshot.HeaderPrimaryFp ?? "", current.HeaderPrimaryFp ?? "", StringComparison.Ordinal))
                errors.Add("HeaderFooterFingerprint changed"); // 例: ［挿入］ヘッダー／フッター
            if (!exemptFlags.HasFlag(WordValidationExemptFlags.CompatibilityMode)
                && snapshot.CompatibilityMode >= 0 && current.CompatibilityMode >= 0
                && current.CompatibilityMode != snapshot.CompatibilityMode)
                errors.Add($"CompatibilityMode changed {snapshot.CompatibilityMode}->{current.CompatibilityMode}"); // 例: ［ファイル］情報の変換
            if (!exemptFlags.HasFlag(WordValidationExemptFlags.Watermark)
                && !string.Equals(snapshot.WatermarkFingerprint ?? "None", current.WatermarkFingerprint ?? "None", StringComparison.Ordinal))
                errors.Add($"WatermarkFingerprint changed {snapshot.WatermarkFingerprint}->{current.WatermarkFingerprint}");
            if (!exemptFlags.HasFlag(WordValidationExemptFlags.PageBorder)
                && !string.Equals(snapshot.PageBorderFingerprint ?? "", current.PageBorderFingerprint ?? "", StringComparison.Ordinal))
                errors.Add($"PageBorderFingerprint changed {snapshot.PageBorderFingerprint}->{current.PageBorderFingerprint}");
            return errors;
        }

        private static SnapshotData CaptureProjectDocument(int projectId, int groupId)
        {
            WordApp app = null;
            try
            {
                try { app = (WordApp)Marshal.GetActiveObject("Word.Application"); }
                catch { return null; }
                if (app == null) return null;
                string path = ResolveSnapshotProjectPath(projectId, groupId);
                if (string.IsNullOrEmpty(path)) return null;
                WordDoc doc = FindOpenDocumentForSnapshot(app, path);
                if (doc == null)
                {
                    try { doc = app.Documents.Open(path, ReadOnly: true, Visible: false); }
                    catch { return null; }
                }
                try { return BuildSnapshotFromDocument(doc); }
                finally
                {
                    if (doc != null)
                    {
                        try { Marshal.ReleaseComObject(doc); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[WordSnapshotChecker] Capture: " + ex.Message);
                return null;
            }
        }

        private static SnapshotData BuildSnapshotFromDocument(WordDoc doc)
        {
            var data = new SnapshotData { FullName = doc.FullName ?? "" };
            try { data.Sections = doc.Sections.Count; } catch { }
            try { data.BodyTextLength = doc.Content.Text?.Length ?? 0; } catch { }
            try { data.InlineShapes = doc.InlineShapes.Count; } catch { }
            try { data.FloatingShapes = doc.Shapes.Count; } catch { }
            try { data.Comments = doc.Comments.Count; } catch { }
            try { data.Tables = doc.Tables.Count; } catch { }
            try { data.CompatibilityMode = (int)doc.CompatibilityMode; } catch { data.CompatibilityMode = -1; }
            data.HeaderPrimaryFp = GetHeaderFingerprint(doc);
            data.WatermarkFingerprint = WordWatermarkInspection.GetWatermarkFingerprintFromDocument(doc);
            data.PageBorderFingerprint = WordWatermarkInspection.GetPageBorderFingerprint(doc);
            return data;
        }

        private static string GetHeaderFingerprint(WordDoc doc)
        {
            try
            {
                if (doc.Sections.Count < 1) return "";
                var hdr = doc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                if (hdr?.Range == null) return "";
                string xml = hdr.Range.WordOpenXML ?? "";
                if (xml.IndexOf("ED7D31", StringComparison.OrdinalIgnoreCase) >= 0) return "ED7D31";
                if (xml.IndexOf("E97132", StringComparison.OrdinalIgnoreCase) >= 0) return "E97132";
                if (xml.IndexOf("accent2", StringComparison.OrdinalIgnoreCase) >= 0) return "accent2";
                return xml.Length > 80 ? xml.Substring(0, 80) : xml;
            }
            catch { return ""; }
        }

        private static WordDoc FindOpenDocumentForSnapshot(WordApp app, string path)
        {
            if (app == null || string.IsNullOrEmpty(path)) return null;
            string pathLower = Path.GetFullPath(path).ToLowerInvariant();
            for (int i = app.Documents.Count; i >= 1; i--)
            {
                WordDoc doc = null;
                try
                {
                    doc = app.Documents[i];
                    string fn = doc.FullName ?? "";
                    try { fn = Path.GetFullPath(fn).ToLowerInvariant(); } catch { fn = fn.ToLowerInvariant(); }
                    if (fn == pathLower)
                        return doc;
                }
                catch { }
                if (doc != null)
                {
                    try { Marshal.ReleaseComObject(doc); } catch { }
                }
            }
            return null;
        }

        private static string ResolveSnapshotProjectPath(int projectId, int groupId)
        {
            string basePath = @"C:\MOSTest\Word365";
            string workingFolder = Path.Combine(basePath, $"Tab{groupId}");
            string[] possibleNames = (groupId == 1 && projectId == 7)
                ? new[] { $"Project{projectId}.doc", $"project{projectId}.doc" }
                : new[] { $"Project{projectId}.docx", $"Project{projectId}.doc", $"project{projectId}.docx", $"project{projectId}.doc" };
            foreach (var fileName in possibleNames)
            {
                string fullPath = Path.Combine(workingFolder, fileName);
                if (File.Exists(fullPath))
                    return fullPath;
            }
            return null;
        }

        private static SnapshotData LoadSnapshot()
        {
            string path = LogReader.GetSnapshotFilePath();
            if (!File.Exists(path))
                return null;
            try
            {
                var data = new SnapshotData();
                foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                        continue;
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string key = line.Substring(0, eq).Trim();
                    string val = line.Substring(eq + 1).Trim();
                    switch (key)
                    {
                        case "ProjectId": int.TryParse(val, out int p); data.ProjectId = p; break;
                        case "TaskId": int.TryParse(val, out int t); data.TaskId = t; break;
                        case "AttemptNo": int.TryParse(val, out int a); data.AttemptNo = a; break;
                        case "Sections": int.TryParse(val, out int s); data.Sections = s; break;
                        case "BodyTextLength": int.TryParse(val, out int bl); data.BodyTextLength = bl; break;
                        case "InlineShapes": int.TryParse(val, out int ins); data.InlineShapes = ins; break;
                        case "FloatingShapes": int.TryParse(val, out int fs); data.FloatingShapes = fs; break;
                        case "Tables": int.TryParse(val, out int tb); data.Tables = tb; break;
                        case "Comments": int.TryParse(val, out int cm); data.Comments = cm; break;
                        case "HeaderPrimaryFp": data.HeaderPrimaryFp = val; break;
                        case "CompatibilityMode": int.TryParse(val, out int c); data.CompatibilityMode = c; break;
                        case "WatermarkFingerprint": data.WatermarkFingerprint = val; break;
                        case "PageBorderFingerprint": data.PageBorderFingerprint = val; break;
                    }
                }
                if (string.IsNullOrEmpty(data.WatermarkFingerprint))
                    data.WatermarkFingerprint = "None";
                return data;
            }
            catch { return null; }
        }
    }

    /// <summary>
    /// 採点前の破壊的操作ゲート。WordChecker（COM 採点）の前に必ず通す。
    /// </summary>
    public static class WordGradingGate
    {
        /// <summary>
        /// ① mos_word_destructive_errors.log（タスク切替時の VSTO 記録 or 過去の採点③）
        /// ② mos_word_log.txt の [Op]（許可外: 例 7-3 の［ホーム］切り取り＝Cut）
        /// ③ スナップショット差分（例: ［レイアウト］区切りで SectionsCount 変化）→ 失敗時に destructive へ追記
        /// </summary>
        public static bool TryPass(int groupId, int projectId, int taskId, int attemptNo, out string failReason)
        {
            failReason = null;
            if (LogReader.HasLoggedDestructiveError(projectId, taskId, attemptNo))
            {
                failReason = "logged destructive error";
                return false;
            }
            var allowedOps = WordTaskValidationConfig.GetAllowedOperationTypes(projectId, taskId);
            if (LogReader.HasDisallowedOperations(projectId, taskId, attemptNo, allowedOps))
            {
                failReason = "disallowed operation in log";
                return false;
            }
            var exempt = WordTaskValidationConfig.GetExemptFlags(projectId, taskId);
            var errors = WordSnapshotChecker.CompareAndGetErrors(groupId, projectId, taskId, attemptNo, exempt);
            if (errors != null && errors.Count > 0)
            {
                LogReader.AppendDestructiveErrors(projectId, taskId, attemptNo, errors);
                failReason = string.Join(" | ", errors);
                return false;
            }
            return true;
        }
    }
}
