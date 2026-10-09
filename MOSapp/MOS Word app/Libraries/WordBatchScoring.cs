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
        private const int ReadyPollIntervalMs = 30;
        private const int DocumentReadyTimeoutMs = 3000;
        private const int DocumentsClosedTimeoutMs = 2000;
        private const int WordStartupTimeoutMs = 3000;
        private const string LastTaskFlushAttemptedKey = "MOS.Word.LastTaskEvidenceFlush.Attempted";
        private const string LastTaskFlushCompletedKey = "MOS.Word.LastTaskEvidenceFlush.Completed";
        private const string LastTaskFlushElapsedKey = "MOS.Word.LastTaskEvidenceFlush.ElapsedMs";
        private const string LastTaskFlushProjectKey = "MOS.Word.LastTaskEvidenceFlush.ProjectId";
        private const string LastTaskFlushTaskKey = "MOS.Word.LastTaskEvidenceFlush.TaskId";
        private const string LastTaskFlushAttemptKey = "MOS.Word.LastTaskEvidenceFlush.AttemptNo";
        private static int _batchVstoFlushFallbackCount;

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
        /// 問題文 JSON を読み込み、指定プロジェクト（未指定時は全件）を採点して ScoreResultStore に保存する。
        /// progress は (メッセージ, 完了プロジェクト数, 対象プロジェクト数)。
        /// </summary>
        public static void ScoreAllProjects(int groupId, ISet<int> projectIds = null, Action<string, int, int> progress = null)
        {
            var projectData = LoadProjectData();
            if (projectData?.Projects == null || projectData.Projects.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[WordBatchScoring] No project data loaded");
                return;
            }

            var projectsToScore = projectData.Projects
                .Where(project => project != null)
                .Where(project => projectIds == null || projectIds.Count == 0 || projectIds.Contains(project.ProjectId))
                .Where(project => (project.Tasks?.Count ?? 0) > 0)
                .OrderBy(project => project.ProjectId)
                .ToList();

            WordGradingPerf.BeginSession("Word batch scoring");
            var totalSw = Stopwatch.StartNew();
            _batchVstoFlushFallbackCount = 0;
            LogLastTaskEvidenceFlush();
            LogReader.BeginBatchScoringLogCache();
            ScoreResultStore.ClearGroup(groupId);
            ReportProgress(progress, "採点の準備をしています...", 0, projectsToScore.Count);

            WordApp wordApp = null;
            try
            {
                var connectSw = Stopwatch.StartNew();
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    wordApp = new WordApp();
                    try { wordApp.Visible = true; } catch { }
                    if (!WaitUntilWordApplicationReady(wordApp))
                    {
                        AppendScoringErrorLog(
                            $"ScoreAllProjects group={groupId} Word startup",
                            new TimeoutException("Word application was not ready"));
                    }
                }
                WordGradingPerf.Log("ScoreAllProjects.Connect", connectSw.ElapsedMilliseconds);

                if (wordApp != null)
                {
                    try { wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone; } catch { }
                    try { wordApp.Visible = true; } catch { }
                    WordWindowLayoutHelper.PositionWordForBatchScoring(wordApp);
                    SaveAllOpenDocuments(wordApp);
                }

                int completedProjects = 0;
                foreach (var project in projectsToScore)
                {
                    int taskCount = project.Tasks?.Count ?? 0;
                    System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] Scoring project {project.ProjectId} ({taskCount} tasks)");
                    ReportProgress(progress, $"プロジェクト {project.ProjectId} を採点中", completedProjects, projectsToScore.Count);
                    var projectSw = Stopwatch.StartNew();

                    try
                    {
                        var closeSw = Stopwatch.StartNew();
                        CloseAllOpenDocumentsForBatch(wordApp);
                        bool closed = WaitUntilDocumentsClosed(wordApp);
                        WordGradingPerf.Log(
                            "OpenProjectDocument.CloseDocuments",
                            closeSw.ElapsedMilliseconds,
                            $"P{project.ProjectId} closed={closed}");
                        if (!closed)
                        {
                            AppendScoringErrorLog(
                                $"ScoreAllProjects group={groupId} project={project.ProjectId} close",
                                new TimeoutException("Word documents were not closed"));
                        }

                        var openSw = Stopwatch.StartNew();
                        bool opened = OpenProjectDocument(wordApp, project.ProjectId, groupId);
                        WordGradingPerf.Log(
                            "OpenProjectDocument.Open",
                            openSw.ElapsedMilliseconds,
                            $"P{project.ProjectId} opened={opened}");
                        if (!opened)
                        {
                            for (int t = 1; t <= taskCount; t++)
                            {
                                ScoreResultStore.RecordResult(
                                    groupId, project.ProjectId, t, false,
                                    WordScoreExplanation.UnavailableText);
                            }
                            continue;
                        }

                        var results = ScoreProject(groupId, project.ProjectId, taskCount, batchMode: true);
                        for (int i = 0; i < taskCount; i++)
                        {
                            bool passed = i < results.Count && results[i].Passed;
                            string failReason = i < results.Count ? results[i].FailReason : WordScoreExplanation.UnavailableText;
                            ScoreResultStore.RecordResult(groupId, project.ProjectId, i + 1, passed, failReason);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[WordBatchScoring] Project {project.ProjectId} failed: {ex.Message}");
                        AppendScoringErrorLog(
                            $"ScoreAllProjects group={groupId} project={project.ProjectId}", ex);
                        for (int t = 1; t <= taskCount; t++)
                        {
                            ScoreResultStore.RecordResult(
                                groupId, project.ProjectId, t, false,
                                WordScoreExplanation.UnavailableText);
                        }
                    }
                    finally
                    {
                        var closeSw = Stopwatch.StartNew();
                        CloseAllOpenDocumentsForBatch(wordApp);
                        WaitUntilDocumentsClosed(wordApp);
                        WordGradingPerf.Log(
                            "OpenProjectDocument.CloseDocuments",
                            closeSw.ElapsedMilliseconds,
                            $"P{project.ProjectId} after-score");
                        completedProjects++;
                        WordGradingPerf.Log(
                            "ScoreAllProjects.Project",
                            projectSw.ElapsedMilliseconds,
                            $"P{project.ProjectId}");
                        ReportProgress(
                            progress,
                            $"プロジェクト {project.ProjectId} の採点が完了しました",
                            completedProjects,
                            projectsToScore.Count);
                    }
                }
            }
            finally
            {
                WordGradingPerf.Log("ScoreAllProjects.Total", totalSw.ElapsedMilliseconds, $"group={groupId}");
                WordGradingPerf.Log(
                    "ScoreAllProjects.VstoFlushFallback",
                    _batchVstoFlushFallbackCount,
                    _batchVstoFlushFallbackCount > 0 ? "fallback=True" : "fallback=False");
                WordGradingPerf.EndSession();
                LogReader.EndBatchScoringLogCache();
                ClearLastTaskEvidenceFlush();
                if (wordApp != null)
                {
                    try { Marshal.ReleaseComObject(wordApp); } catch { }
                }
            }

            System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] Done scoring group {groupId}");
        }

        private static void ReportProgress(Action<string, int, int> progress, string message, int completed, int total)
        {
            if (progress == null)
                return;
            try
            {
                progress(message, completed, total);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] Progress callback: {ex.Message}");
            }
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
                    try { wordApp.Visible = true; } catch { }
                    if (!WaitUntilWordApplicationReady(wordApp))
                    {
                        AppendScoringErrorLog(
                            $"ScoreSingleTask group={groupId} project={projectId} task={taskId} Word startup",
                            new TimeoutException("Word application was not ready"));
                        return null;
                    }
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

                if (!WaitUntilDocumentReady(wordApp, ResolveProjectFilePath(projectId, groupId)))
                {
                    AppendScoringErrorLog(
                        $"ScoreSingleTask group={groupId} project={projectId} task={taskId} ready",
                        new TimeoutException("Word document was not ready"));
                    return null;
                }
                int attemptNo = WordTaskAttemptRegistry.GetAttempt(projectId, taskId);
                WordScoreExplanation.ClearCheckerReason();
                if (!WordGradingGate.TryPass(groupId, projectId, taskId, attemptNo, out string gateReason))
                {
                    string failReason = WordScoreExplanation.ResolveFailReason(
                        passed: false, gateFailed: true, gateInternalReason: gateReason);
                    ScoreResultStore.RecordResult(groupId, projectId, taskId, false, failReason);
                    return false;
                }
                bool? result = InvokeCheckTask(groupId, projectId, taskId);
                if (result.HasValue)
                {
                    string failReason = WordScoreExplanation.ResolveFailReason(result.Value);
                    ScoreResultStore.RecordResult(groupId, projectId, taskId, result.Value, failReason);
                }
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
            LogReader.RequestVstoEvidenceFlush();
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
                string jsonPath = MOS_Word_app.WordDataPathHelper.FindProblemJson("MOS模擬アプリ問題文一覧_Word.json");
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

        /// <summary>
        /// レビュー遷移で Word 文書を閉じる直前に、表示中タスクの VSTO 証跡を1回確定する。
        /// ack が指定タスクと一致したセッションだけ、一括採点中のタスク別フラッシュを省略する。
        /// </summary>
        public static void PrepareBeforeClosingDocumentsForBatchScoring(int projectId, int taskId, int attemptNo)
        {
            var sw = Stopwatch.StartNew();
            VstoEvidenceFlushAck ack = null;
            try
            {
                ack = LogReader.TryRequestVstoEvidenceFlush(projectId, taskId, attemptNo);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[WordBatchScoring] Last-task evidence flush: " + ex.Message);
            }

            bool completed = ack != null
                && ack.Completed
                && ack.ProjectId == projectId
                && ack.TaskId == taskId
                && ack.AttemptNo == attemptNo
                && projectId > 0
                && taskId > 0;
            AppDomain.CurrentDomain.SetData(LastTaskFlushAttemptedKey, true);
            AppDomain.CurrentDomain.SetData(LastTaskFlushCompletedKey, completed);
            AppDomain.CurrentDomain.SetData(LastTaskFlushElapsedKey, sw.ElapsedMilliseconds);
            AppDomain.CurrentDomain.SetData(LastTaskFlushProjectKey, projectId);
            AppDomain.CurrentDomain.SetData(LastTaskFlushTaskKey, taskId);
            AppDomain.CurrentDomain.SetData(LastTaskFlushAttemptKey, attemptNo);
        }

        private static void RequestVstoEvidenceFlushIfNeeded(int groupId, int projectId, int taskNum, bool batchMode)
        {
            if (batchMode && IsLastTaskEvidenceFlushCompleted())
                return;
            if (!batchMode || VSTOCheckerHelper.RequiresVstoEvidenceFlush(groupId, projectId, taskNum))
            {
                if (batchMode)
                    _batchVstoFlushFallbackCount++;
                var flushSw = Stopwatch.StartNew();
                int attemptNo = WordTaskAttemptRegistry.GetAttempt(projectId, taskNum);
                LogReader.TryRequestVstoEvidenceFlush(projectId, taskNum, attemptNo);
                WordGradingPerf.Log(
                    "VstoEvidenceFlush",
                    flushSw.ElapsedMilliseconds,
                    $"P{projectId}-{taskNum} fallback={batchMode}");
            }
        }

        private static void LogLastTaskEvidenceFlush()
        {
            bool attempted = AppDomain.CurrentDomain.GetData(LastTaskFlushAttemptedKey) as bool? == true;
            bool completed = AppDomain.CurrentDomain.GetData(LastTaskFlushCompletedKey) as bool? == true;
            int projectId = ReadAppDomainInt(LastTaskFlushProjectKey);
            int taskId = ReadAppDomainInt(LastTaskFlushTaskKey);
            int attemptNo = ReadAppDomainInt(LastTaskFlushAttemptKey);
            long elapsedMs = 0;
            object elapsed = AppDomain.CurrentDomain.GetData(LastTaskFlushElapsedKey);
            if (elapsed is long value)
                elapsedMs = value;
            WordGradingPerf.Log(
                "ScoreAllProjects.LastTaskEvidenceFlush",
                elapsedMs,
                $"attempted={attempted} completed={completed} task={projectId}-{taskId} attempt={attemptNo} ack={completed}");
        }

        private static bool IsLastTaskEvidenceFlushCompleted()
        {
            return AppDomain.CurrentDomain.GetData(LastTaskFlushCompletedKey) as bool? == true
                && ReadAppDomainInt(LastTaskFlushProjectKey) > 0
                && ReadAppDomainInt(LastTaskFlushTaskKey) > 0;
        }

        private static int ReadAppDomainInt(string key)
        {
            object value = AppDomain.CurrentDomain.GetData(key);
            return value is int number ? number : 0;
        }

        private static void ClearLastTaskEvidenceFlush()
        {
            AppDomain.CurrentDomain.SetData(LastTaskFlushAttemptedKey, null);
            AppDomain.CurrentDomain.SetData(LastTaskFlushCompletedKey, null);
            AppDomain.CurrentDomain.SetData(LastTaskFlushElapsedKey, null);
            AppDomain.CurrentDomain.SetData(LastTaskFlushProjectKey, null);
            AppDomain.CurrentDomain.SetData(LastTaskFlushTaskKey, null);
            AppDomain.CurrentDomain.SetData(LastTaskFlushAttemptKey, null);
            _batchVstoFlushFallbackCount = 0;
        }

        private struct TaskScoreResult
        {
            public bool Passed;
            public string FailReason;
        }

        private static List<TaskScoreResult> ScoreProject(int groupId, int projectId, int taskCount, bool batchMode = false)
        {
            var results = new List<TaskScoreResult>();
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string dllPath = Path.Combine(baseDir, "Dlls", $"WordChecker{groupId}_{projectId}.dll");

            if (!File.Exists(dllPath))
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] DLL not found: {dllPath}");
                for (int i = 0; i < taskCount; i++)
                    results.Add(new TaskScoreResult { Passed = false, FailReason = WordScoreExplanation.UnavailableText });
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
                        results.Add(new TaskScoreResult { Passed = false, FailReason = WordScoreExplanation.UnavailableText });
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
                        results.Add(new TaskScoreResult { Passed = false, FailReason = WordScoreExplanation.UnavailableText });
                        continue;
                    }

                    var taskSw = Stopwatch.StartNew();
                    try
                    {
                        WordScoreExplanation.ClearCheckerReason();
                        int attemptNo = WordTaskAttemptRegistry.GetAttempt(projectId, taskNum);
                        var gateSw = Stopwatch.StartNew();
                        bool gatePassed = WordGradingGate.TryPass(groupId, projectId, taskNum, attemptNo, out string gateReason);
                        WordGradingPerf.Log(
                            "GradeTask.WordGradingGate",
                            gateSw.ElapsedMilliseconds,
                            $"P{projectId}-{taskNum} passed={gatePassed}");
                        if (!gatePassed)
                        {
                            string failReason = WordScoreExplanation.ResolveFailReason(
                                passed: false, gateFailed: true, gateInternalReason: gateReason);
                            results.Add(new TaskScoreResult { Passed = false, FailReason = failReason });
                            System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] P{projectId} T{taskNum}: fail (grading gate)");
                            WordGradingPerf.Log("GradeTask.total", taskSw.ElapsedMilliseconds, $"P{projectId}-{taskNum} gate");
                            continue;
                        }
                        RequestVstoEvidenceFlushIfNeeded(groupId, projectId, taskNum, batchMode);
                        var checkerSw = Stopwatch.StartNew();
                        bool taskResult = (bool)method.Invoke(checkerInstance, null);
                        WordGradingPerf.Log(
                            "GradeTask.CheckerInvoke",
                            checkerSw.ElapsedMilliseconds,
                            $"P{projectId}-{taskNum} passed={taskResult}");
                        string reason = WordScoreExplanation.ResolveFailReason(taskResult);
                        results.Add(new TaskScoreResult { Passed = taskResult, FailReason = reason });
                        WordGradingPerf.Log("GradeTask.total", taskSw.ElapsedMilliseconds, $"P{projectId}-{taskNum}");
                        System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] P{projectId} T{taskNum}: {(taskResult ? "pass" : "fail")}");
                    }
                    catch (Exception exTask)
                    {
                        System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] P{projectId} T{taskNum} error: {exTask.Message}");
                        WordScoreExplanation.ClearCheckerReason();
                        results.Add(new TaskScoreResult
                        {
                            Passed = false,
                            FailReason = WordScoreExplanation.ResolveFailReason(passed: false, unavailable: true)
                        });
                        WordGradingPerf.Log("GradeTask.total", taskSw.ElapsedMilliseconds, $"P{projectId}-{taskNum} error");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] ScoreProject error: {ex.Message}");
                while (results.Count < taskCount)
                    results.Add(new TaskScoreResult { Passed = false, FailReason = WordScoreExplanation.UnavailableText });
            }

            return results;
        }

        /// <summary>
        /// 一括採点: 開いている文書をすべて閉じる（1プロジェクト1文書の安定運用）。
        /// </summary>
        private static void CloseAllOpenDocumentsForBatch(WordApp wordApp)
        {
            if (wordApp == null)
                return;

            try
            {
                if (wordApp.Documents.Count == 0)
                    return;

                wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            }
            catch { }

            const int maxAttempts = 50;
            for (int attempt = 0; attempt < maxAttempts && wordApp.Documents.Count > 0; attempt++)
            {
                int countBefore = wordApp.Documents.Count;
                if (!TryCloseBatchDocumentAt(wordApp, countBefore))
                {
                    if (countBefore > 1 && TryCloseBatchDocumentAt(wordApp, 1))
                        continue;

                    System.Diagnostics.Debug.WriteLine(
                        $"[WordBatchScoring] Document close made no progress (remaining={wordApp.Documents.Count})");
                    break;
                }
            }
        }

        private static bool TryCloseBatchDocumentAt(WordApp wordApp, int index)
        {
            if (wordApp == null || index < 1 || index > wordApp.Documents.Count)
                return false;

            int countBefore = wordApp.Documents.Count;
            WordDoc doc = null;
            try
            {
                doc = wordApp.Documents[index];
                doc.Close(SaveChanges: false);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] TryCloseBatchDocumentAt: {ex.Message}");
            }
            finally
            {
                if (doc != null)
                {
                    try { Marshal.ReleaseComObject(doc); } catch { }
                }
            }

            return wordApp.Documents.Count < countBefore;
        }

        /// <summary>起動時に前回までの採点例外ログを消す。試験のリセットでは消さない。</summary>
        public static void ClearScoringErrorLog()
        {
            try
            {
                string logPath = Path.Combine(Path.GetTempPath(), "mos_word_scoring_errors.log");
                if (File.Exists(logPath))
                    File.Delete(logPath);
            }
            catch
            {
                // ignore log failure
            }
        }

        private static void AppendScoringErrorLog(string context, Exception ex)
        {
            try
            {
                string logPath = Path.Combine(Path.GetTempPath(), "mos_word_scoring_errors.log");
                var sb = new StringBuilder();
                sb.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}");
                sb.AppendLine($"Message: {ex.Message}");
                if (ex is COMException comEx)
                    sb.AppendLine($"HResult: 0x{comEx.ErrorCode:X8}");
                else
                    sb.AppendLine($"HResult: 0x{ex.HResult:X8}");
                sb.AppendLine(ex.ToString());
                sb.AppendLine("---");
                File.AppendAllText(logPath, sb.ToString(), Encoding.UTF8);
            }
            catch
            {
                // ignore log failure
            }
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
                            WordWindowLayoutHelper.PositionWordForBatchScoring(wordApp);
                            return WaitUntilOpenedDocumentReady(wordApp, filePath, groupId, projectId);
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
                WordWindowLayoutHelper.PositionWordForBatchScoring(wordApp);
                return WaitUntilOpenedDocumentReady(wordApp, filePath, groupId, projectId);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] OpenProjectDocument error: {ex.Message}");
                return false;
            }
        }

        private static string ResolveProjectFilePath(int projectId, int groupId)
        {
            try
            {
                return MOS_Word_app.WordDataPathHelper.EnsureWorkingFile(groupId, projectId);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordBatchScoring] Resolve project file failed: {ex.Message}");
                return null;
            }
        }

        private static bool WaitUntilWordApplicationReady(WordApp wordApp)
        {
            var wait = Stopwatch.StartNew();
            while (wait.ElapsedMilliseconds < WordStartupTimeoutMs)
            {
                if (IsWordApplicationReady(wordApp))
                    return true;
                Thread.Sleep(ReadyPollIntervalMs);
            }
            return IsWordApplicationReady(wordApp);
        }

        private static bool IsWordApplicationReady(WordApp wordApp)
        {
            if (wordApp == null)
                return false;
            try
            {
                bool visible = wordApp.Visible;
                int count = wordApp.Documents.Count;
                return visible && count >= 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool WaitUntilDocumentsClosed(WordApp wordApp)
        {
            var wait = Stopwatch.StartNew();
            while (wait.ElapsedMilliseconds < DocumentsClosedTimeoutMs)
            {
                if (AreDocumentsClosed(wordApp))
                    return true;
                Thread.Sleep(ReadyPollIntervalMs);
            }
            return AreDocumentsClosed(wordApp);
        }

        private static bool AreDocumentsClosed(WordApp wordApp)
        {
            if (wordApp == null)
                return true;
            try
            {
                return wordApp.Documents.Count == 0;
            }
            catch (COMException)
            {
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool WaitUntilOpenedDocumentReady(WordApp wordApp, string filePath, int groupId, int projectId)
        {
            var readySw = Stopwatch.StartNew();
            bool ready = WaitUntilDocumentReady(wordApp, filePath);
            WordGradingPerf.Log(
                "OpenProjectDocument.DocumentReady",
                readySw.ElapsedMilliseconds,
                $"P{projectId} ready={ready}");
            if (!ready)
            {
                AppendScoringErrorLog(
                    $"OpenProjectDocument timeout group={groupId} project={projectId} path={filePath}",
                    new TimeoutException("Word document was not ready"));
            }
            return ready;
        }

        private static bool WaitUntilDocumentReady(WordApp wordApp, string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            var wait = Stopwatch.StartNew();
            while (wait.ElapsedMilliseconds < DocumentReadyTimeoutMs)
            {
                if (IsDocumentReady(wordApp, filePath))
                    return true;
                Thread.Sleep(ReadyPollIntervalMs);
            }
            return IsDocumentReady(wordApp, filePath);
        }

        private static bool IsDocumentReady(WordApp wordApp, string filePath)
        {
            WordDoc active = null;
            Microsoft.Office.Interop.Word.Window window = null;
            Microsoft.Office.Interop.Word.Range content = null;
            try
            {
                if (wordApp == null || string.IsNullOrEmpty(filePath))
                    return false;
                if (!ContainsDocument(wordApp, filePath))
                    return false;

                active = wordApp.ActiveDocument;
                if (active == null || !DocumentPathsEqual(active.FullName, filePath))
                    return false;

                window = active.ActiveWindow;
                if (window == null || !window.Visible || !wordApp.Visible)
                    return false;

                content = active.Content;
                if (content == null)
                    return false;
                int start = content.Start;
                return start >= 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (content != null)
                {
                    try { Marshal.ReleaseComObject(content); } catch { }
                }
                if (window != null)
                {
                    try { Marshal.ReleaseComObject(window); } catch { }
                }
                if (active != null)
                {
                    try { Marshal.ReleaseComObject(active); } catch { }
                }
            }
        }

        private static bool ContainsDocument(WordApp wordApp, string filePath)
        {
            int count;
            try { count = wordApp.Documents.Count; }
            catch { return false; }

            for (int i = 1; i <= count; i++)
            {
                WordDoc doc = null;
                try
                {
                    doc = wordApp.Documents[i];
                    if (DocumentPathsEqual(doc.FullName, filePath))
                        return true;
                }
                catch { }
                finally
                {
                    if (doc != null)
                    {
                        try { Marshal.ReleaseComObject(doc); } catch { }
                    }
                }
            }
            return false;
        }

        private static bool DocumentPathsEqual(string left, string right)
        {
            return string.Equals(NormalizeDocumentPath(left), NormalizeDocumentPath(right), StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeDocumentPath(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return string.Empty;
            try { return Path.GetFullPath(filePath); }
            catch { return filePath; }
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
                    // 9-5: 変更履歴の承認で表・図形・ヘッダー等が確定し得る。破壊判定は可視本文長に限る。
                    flags |= WordValidationExemptFlags.SectionsCount
                        | WordValidationExemptFlags.BodyTextLength
                        | WordValidationExemptFlags.InlineShapesCount
                        | WordValidationExemptFlags.FloatingShapesCount
                        | WordValidationExemptFlags.TablesCount
                        | WordValidationExemptFlags.CommentsCount
                        | WordValidationExemptFlags.HeaderFooterFingerprint
                        | WordValidationExemptFlags.PageBorder
                        | WordValidationExemptFlags.Watermark
                        | WordValidationExemptFlags.CompatibilityMode;
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
                // VSTO が切り取り選択を診断用に記録（ユーザー操作ではない）
                allowed.Add("CutSelection");
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
            public int FootnoteReferenceCount { get; set; }
            public int VisibleBodyTextLength { get; set; } = -1;
        }

        public static List<string> CompareAndGetErrors(int groupId, int projectId, int taskId, int attemptNo, WordValidationExemptFlags exemptFlags)
        {
            var sw = Stopwatch.StartNew();
            try
            {
                return CompareAndGetErrorsCore(groupId, projectId, taskId, attemptNo, exemptFlags);
            }
            finally
            {
                WordGradingPerf.Log(
                    "WordSnapshotChecker.CompareAndGetErrors",
                    sw.ElapsedMilliseconds,
                    $"P{projectId} T{taskId}");
            }
        }

        private static List<string> CompareAndGetErrorsCore(int groupId, int projectId, int taskId, int attemptNo, WordValidationExemptFlags exemptFlags)
        {
            var errors = new List<string>();
            var snapshot = LoadSnapshot();
            if (snapshot == null)
                return errors;
            if (snapshot.ProjectId != projectId || snapshot.TaskId != taskId || snapshot.AttemptNo != attemptNo)
                return errors;
            string expectedPath = ResolveSnapshotProjectPath(projectId, groupId);
            if (!SnapshotPathMatches(snapshot.FullName, expectedPath))
            {
                WordGradingPerf.Log(
                    "WordSnapshotChecker.SkippedPathMismatch",
                    0,
                    $"P{projectId} T{taskId} snapshot={snapshot.FullName} expected={expectedPath}");
                return errors;
            }

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
            if (projectId == 8 && taskId == 3
                && current.FootnoteReferenceCount != snapshot.FootnoteReferenceCount)
                errors.Add($"FootnoteReferenceCount changed {snapshot.FootnoteReferenceCount}->{current.FootnoteReferenceCount}");
            if (projectId == 9 && (taskId == 3 || taskId == 4 || taskId == 5)
                && snapshot.VisibleBodyTextLength >= 0 && current.VisibleBodyTextLength >= 0
                && current.VisibleBodyTextLength != snapshot.VisibleBodyTextLength)
                errors.Add($"VisibleBodyTextLength changed {snapshot.VisibleBodyTextLength}->{current.VisibleBodyTextLength}");
            return errors;
        }

        /// <summary>
        /// その場採点の開始時、スナップショットが一致するタスクだけ開始時との差分を一度記録する。
        /// 一致しないときは文書を読み直さない。
        /// </summary>
        [ThreadStatic] static bool _reuseActive;
        [ThreadStatic] static bool _reuseCaptured;
        [ThreadStatic] static int _reuseGroupId;
        [ThreadStatic] static int _reuseProjectId;
        [ThreadStatic] static SnapshotData _reuseCurrent;

        /// <summary>その場採点中は、今の文書の現在状態を1回だけ読む。</summary>
        public static void BeginReuseCurrentDocument(int groupId, int projectId)
        {
            _reuseActive = true;
            _reuseCaptured = false;
            _reuseGroupId = groupId;
            _reuseProjectId = projectId;
            _reuseCurrent = null;
        }

        public static void EndReuseCurrentDocument()
        {
            _reuseActive = false;
            _reuseCaptured = false;
            _reuseCurrent = null;
        }

        public static void LogMatchingBaselineDiffOnce(int groupId, int projectId, int taskId, int attemptNo)
        {
            var exempt = WordTaskValidationConfig.GetExemptFlags(projectId, taskId);
            var errors = CompareAndGetErrors(groupId, projectId, taskId, attemptNo, exempt);
            if (errors != null && errors.Count > 0)
                LogReader.AppendDestructiveErrors(projectId, taskId, attemptNo, errors);
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
                if (_reuseActive && _reuseCaptured && groupId == _reuseGroupId && projectId == _reuseProjectId)
                    return _reuseCurrent;

                WordDoc doc = FindOpenDocumentForSnapshot(app, path);
                if (doc == null)
                {
                    try { doc = app.Documents.Open(path, ReadOnly: true, Visible: false); }
                    catch { return RememberReuse(groupId, projectId, null); }
                }
                try { return RememberReuse(groupId, projectId, BuildSnapshotFromDocument(doc)); }
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

        static SnapshotData RememberReuse(int groupId, int projectId, SnapshotData data)
        {
            if (_reuseActive && groupId == _reuseGroupId && projectId == _reuseProjectId)
            {
                _reuseCaptured = true;
                _reuseCurrent = data;
            }
            return data;
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
            data.PageBorderFingerprint = WordWatermarkInspection.GetPageBorderFingerprint(doc);
            try
            {
                string openXml = doc.WordOpenXML;
                data.WatermarkFingerprint = WordWatermarkInspection.GetWatermarkFingerprint(
                    WordWatermarkInspection.NormalizeXml(openXml));
                data.FootnoteReferenceCount = WordFindHelper.CountFootnoteReferencesInXml(openXml);
                data.VisibleBodyTextLength = WordFindHelper.CountVisibleBodyTextLength(openXml);
            }
            catch
            {
                data.FootnoteReferenceCount = 0;
            }
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
            return MOS_Word_app.WordDataPathHelper.FindExistingWorkingFile(groupId, projectId);
        }

        private static bool SnapshotPathMatches(string snapshotPath, string expectedPath)
        {
            if (string.IsNullOrWhiteSpace(snapshotPath) || string.IsNullOrWhiteSpace(expectedPath))
                return false;
            try
            {
                return string.Equals(
                    Path.GetFullPath(snapshotPath),
                    Path.GetFullPath(expectedPath),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(snapshotPath.Trim(), expectedPath.Trim(), StringComparison.OrdinalIgnoreCase);
            }
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
                        case "FullName": data.FullName = val; break;
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
                        case "FootnoteReferenceCount": int.TryParse(val, out int fn); data.FootnoteReferenceCount = fn; break;
                        case "VisibleBodyTextLength": int.TryParse(val, out int visible); data.VisibleBodyTextLength = visible; break;
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
            if (projectId == 2 && taskId == 1
                && LogReader.HasTaskEvidence(2, 1, "CutParagraphSelection"))
            {
                failReason = "CutParagraphSelection detected";
                return false;
            }
            if (projectId == 4 && taskId == 3
                && LogReader.HasTaskEvidence(4, 3, "ReviewDeleteComment")
                && !LogReader.HasTaskEvidence(4, 3, "ReviewResolveComment"))
            {
                failReason = "ReviewDeleteComment without resolve";
                return false;
            }
            if (projectId == 4 && taskId == 3
                && LogReader.HasTaskEvidence(4, 3, "ReviewResolveComment")
                && !LogReader.HasTaskEvidence(4, 3, "ReviewDeleteComment"))
            {
                failReason = "ReviewResolveComment without delete";
                return false;
            }
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
