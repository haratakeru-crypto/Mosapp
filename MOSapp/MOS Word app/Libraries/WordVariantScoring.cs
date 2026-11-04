using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Libraries;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WordDoc = Microsoft.Office.Interop.Word.Document;

namespace MOS_Word_app
{
    /// <summary>
    /// 類題第1回のその場採点と一括採点。採点前に破壊的操作ゲート（セット1）を通す。
    /// 未実装メソッド・DLL なしは IsPassed=null（画面は「－」）。
    /// </summary>
    public static class WordVariantScoring
    {
        public const string SetSuffix = "PV1";
        public const int DefaultVariantSetNo = 1;
        public const string UnimplementedReason = WordScoreExplanation.UnimplementedText;

        public static string GetLibraryName(int groupId, int projectId)
        {
            return $"WordChecker{groupId}_{projectId}_{SetSuffix}";
        }

        public static string GetDllPath(int groupId, int projectId)
        {
            return Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Dlls",
                GetLibraryName(groupId, projectId) + ".dll");
        }

        public static string GetTypeName(int groupId, int projectId)
        {
            return $"Libraries.Group{groupId}.{GetLibraryName(groupId, projectId)}";
        }

        public static string GetMethodName(int groupId, int projectId, int taskId)
        {
            return $"CheckTask_{groupId}_{projectId}_{taskId:D2}";
        }

        public static IReadOnlyList<TaskResult> ScoreProject(int groupId, int projectId, int variantSetNo = DefaultVariantSetNo)
        {
            if (variantSetNo < 1)
                variantSetNo = DefaultVariantSetNo;

            var tasks = VariantProblemWorkbook.GetTasks(projectId);
            var results = new List<TaskResult>(tasks.Count);
            if (tasks.Count == 0)
                return results;

            TryLoadChecker(groupId, projectId, out Type checkerType, out object checkerInstance);

            foreach (VariantTaskText task in tasks)
            {
                int attemptNo = WordTaskAttemptRegistry.GetAttempt(projectId, task.TaskId, variantSetNo);
                if (!WordGradingGate.TryPass(
                    groupId, projectId, task.TaskId, attemptNo, out List<string> gateActions, variantSetNo))
                {
                    results.Add(new TaskResult
                    {
                        TaskNumber = task.TaskId,
                        IsPassed = false,
                        TaskName = $"タスク{task.TaskId}",
                        FailReason = WordScoreExplanation.ResolveFailReason(
                            passed: false, gateFailed: true, gateStudentActions: gateActions)
                    });
                    continue;
                }

                bool? passed = TryScoreTask(checkerType, checkerInstance, groupId, projectId, task.TaskId);
                results.Add(new TaskResult
                {
                    TaskNumber = task.TaskId,
                    IsPassed = passed,
                    TaskName = $"タスク{task.TaskId}",
                    FailReason = passed == false ? WordScoreExplanation.ResolveFailReason(false) : ""
                });
            }

            return results;
        }

        /// <summary>
        /// 類題 P を順に開き、PV1 チェッカーで採点して ScoreResultStore に保存する。
        /// </summary>
        public static void ScoreAllProjects(
            int groupId,
            ISet<int> projectIds = null,
            Action<string, int, int> progress = null,
            int variantSetNo = DefaultVariantSetNo)
        {
            if (variantSetNo < 1)
                variantSetNo = DefaultVariantSetNo;

            var projectIdsToScore = VariantProblemWorkbook.GetProjectIds()
                .Where(id => projectIds == null || projectIds.Count == 0 || projectIds.Contains(id))
                .Where(id => VariantProblemWorkbook.GetTasks(id).Count > 0)
                .OrderBy(id => id)
                .ToList();

            ScoreResultStore.ClearGroup(groupId, variantSetNo);
            ReportProgress(progress, "採点の準備をしています...", 0, projectIdsToScore.Count);

            WordApp wordApp = null;
            try
            {
                wordApp = GetOrCreateWordApp();
                if (wordApp != null)
                {
                    try { wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone; } catch { }
                    try { wordApp.Visible = true; } catch { }
                    WordWindowLayoutHelper.PositionWordForBatchScoring(wordApp);
                }

                int completed = 0;
                foreach (int projectId in projectIdsToScore)
                {
                    var tasks = VariantProblemWorkbook.GetTasks(projectId);
                    ReportProgress(progress, $"プロジェクト {projectId} を採点中", completed, projectIdsToScore.Count);

                    try
                    {
                        CloseOpenDocuments(wordApp);
                        string filePath = WordDataPathHelper.FindVariantWorkingFile(groupId, projectId, variantSetNo);
                        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath) || !OpenDocument(wordApp, filePath))
                        {
                            RecordUnavailable(groupId, projectId, tasks, variantSetNo);
                            continue;
                        }

                        IReadOnlyList<TaskResult> results = ScoreProject(groupId, projectId, variantSetNo);
                        RecordProjectResults(groupId, projectId, tasks, results, variantSetNo);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[WordVariantScoring] Project {projectId} failed: {ex.Message}");
                        RecordUnavailable(groupId, projectId, tasks, variantSetNo);
                    }
                    finally
                    {
                        CloseOpenDocuments(wordApp);
                        completed++;
                        ReportProgress(
                            progress,
                            $"プロジェクト {projectId} の採点が完了しました",
                            completed,
                            projectIdsToScore.Count);
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
        }

        /// <summary>結果画面からの1問再採点。成功時は結果を返し、失敗時は null。</summary>
        public static bool? ScoreSingleTask(
            int groupId,
            int projectId,
            int taskId,
            int variantSetNo = DefaultVariantSetNo)
        {
            if (variantSetNo < 1)
                variantSetNo = DefaultVariantSetNo;

            WordApp wordApp = null;
            try
            {
                wordApp = GetOrCreateWordApp();
                if (wordApp == null)
                    return null;

                try { wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone; } catch { }

                string filePath = WordDataPathHelper.FindVariantWorkingFile(groupId, projectId, variantSetNo);
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                    return null;

                if (!TryActivateDocument(wordApp, filePath) && !OpenDocument(wordApp, filePath))
                    return null;

                int attemptNo = WordTaskAttemptRegistry.GetAttempt(projectId, taskId, variantSetNo);
                if (!WordGradingGate.TryPass(
                    groupId, projectId, taskId, attemptNo, out List<string> gateActions, variantSetNo))
                {
                    string gateFail = WordScoreExplanation.ResolveFailReason(
                        passed: false, gateFailed: true, gateStudentActions: gateActions);
                    ScoreResultStore.RecordResult(groupId, projectId, taskId, false, gateFail, variantSetNo);
                    return false;
                }

                TryLoadChecker(groupId, projectId, out Type checkerType, out object checkerInstance);
                bool? passed = TryScoreTask(checkerType, checkerInstance, groupId, projectId, taskId);
                if (!passed.HasValue)
                {
                    ScoreResultStore.RecordResult(groupId, projectId, taskId, false, UnimplementedReason, variantSetNo);
                    return null;
                }

                string failReason = passed.Value ? "" : WordScoreExplanation.ResolveFailReason(false);
                ScoreResultStore.RecordResult(groupId, projectId, taskId, passed.Value, failReason, variantSetNo);
                return passed;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[WordVariantScoring] ScoreSingleTask P{projectId} T{taskId}: {ex.Message}");
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

        public static bool IsUnimplementedReason(string reason)
        {
            return string.Equals(reason, UnimplementedReason, StringComparison.Ordinal);
        }

        public static bool IsTaskImplemented(int groupId, int projectId, int taskId)
        {
            Type checkerType = GetCheckerType(groupId, projectId);
            if (checkerType == null)
                return false;
            return checkerType.GetMethod(GetMethodName(groupId, projectId, taskId)) != null;
        }

        private static void RecordProjectResults(
            int groupId,
            int projectId,
            IReadOnlyList<VariantTaskText> tasks,
            IReadOnlyList<TaskResult> results,
            int variantSetNo)
        {
            foreach (VariantTaskText task in tasks)
            {
                TaskResult match = results.FirstOrDefault(r => r.TaskNumber == task.TaskId);
                bool? passed = match?.IsPassed;
                if (!passed.HasValue)
                {
                    ScoreResultStore.RecordResult(
                        groupId, projectId, task.TaskId, false, UnimplementedReason, variantSetNo);
                    continue;
                }

                ScoreResultStore.RecordResult(
                    groupId,
                    projectId,
                    task.TaskId,
                    passed.Value,
                    passed.Value ? "" : (match?.FailReason ?? WordScoreExplanation.ResolveFailReason(false)),
                    variantSetNo);
            }
        }

        private static void RecordUnavailable(
            int groupId, int projectId, IReadOnlyList<VariantTaskText> tasks, int variantSetNo)
        {
            foreach (VariantTaskText task in tasks)
            {
                ScoreResultStore.RecordResult(
                    groupId, projectId, task.TaskId, false, WordScoreExplanation.UnavailableText, variantSetNo);
            }
        }

        private static bool? TryScoreTask(
            Type checkerType,
            object checkerInstance,
            int groupId,
            int projectId,
            int taskId)
        {
            if (checkerType == null || checkerInstance == null)
                return null;

            MethodInfo method = checkerType.GetMethod(GetMethodName(groupId, projectId, taskId));
            if (method == null)
                return null;

            try
            {
                if (projectId == 8 && taskId == 5)
                    LogReader.RequestVstoEvidenceFlush();

                WordScoreExplanation.ClearCheckerReason();
                object value = method.Invoke(checkerInstance, null);
                return value is bool flag ? flag : (bool?)null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[WordVariantScoring] タスク{taskId} 呼び出しエラー: {ex.Message}");
                return false;
            }
        }

        private static readonly Dictionary<string, Type> CheckerTypeCache =
            new Dictionary<string, Type>(StringComparer.Ordinal);
        private static readonly object CheckerTypeCacheLock = new object();

        private static Type GetCheckerType(int groupId, int projectId)
        {
            string cacheKey = GetLibraryName(groupId, projectId);
            lock (CheckerTypeCacheLock)
            {
                if (CheckerTypeCache.TryGetValue(cacheKey, out Type cached))
                    return cached;
            }

            Type checkerType = null;
            string dllPath = GetDllPath(groupId, projectId);
            if (!File.Exists(dllPath))
            {
                Debug.WriteLine($"[WordVariantScoring] DLL なし（未実装扱い）: {dllPath}");
            }
            else
            {
                try
                {
                    Assembly assembly = Assembly.LoadFrom(dllPath);
                    string typeName = GetTypeName(groupId, projectId);
                    checkerType = assembly.GetType(typeName);
                    if (checkerType == null)
                        Debug.WriteLine($"[WordVariantScoring] 型なし: {typeName}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[WordVariantScoring] 読み込み失敗: {ex.Message}");
                    checkerType = null;
                }
            }

            lock (CheckerTypeCacheLock)
            {
                CheckerTypeCache[cacheKey] = checkerType;
            }

            return checkerType;
        }

        private static void TryLoadChecker(int groupId, int projectId, out Type checkerType, out object checkerInstance)
        {
            checkerType = GetCheckerType(groupId, projectId);
            checkerInstance = null;
            if (checkerType == null)
                return;

            try
            {
                checkerInstance = Activator.CreateInstance(checkerType);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[WordVariantScoring] インスタンス作成失敗: {ex.Message}");
                checkerType = null;
                checkerInstance = null;
            }
        }

        private static WordApp GetOrCreateWordApp()
        {
            try
            {
                return (WordApp)Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                try
                {
                    var app = new WordApp { Visible = true };
                    return app;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[WordVariantScoring] Word 起動失敗: {ex.Message}");
                    return null;
                }
            }
        }

        private static bool OpenDocument(WordApp wordApp, string filePath)
        {
            if (wordApp == null || string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return false;

            try
            {
                WordDataPathHelper.RemoveZoneIdentifier(filePath);
                WordDoc opened = wordApp.Documents.Open(filePath, ReadOnly: false, Visible: true);
                try { opened.Activate(); } catch { }
                finally
                {
                    if (opened != null)
                    {
                        try { Marshal.ReleaseComObject(opened); } catch { }
                    }
                }
                WordWindowLayoutHelper.PositionWordForBatchScoring(wordApp);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[WordVariantScoring] OpenDocument: {ex.Message}");
                return false;
            }
        }

        private static bool TryActivateDocument(WordApp wordApp, string filePath)
        {
            if (wordApp == null || string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return false;

            try
            {
                string expectedPath = Path.GetFullPath(filePath);
                for (int i = wordApp.Documents.Count; i >= 1; i--)
                {
                    WordDoc doc = null;
                    try
                    {
                        doc = wordApp.Documents[i];
                        string fullName = doc.FullName ?? string.Empty;
                        string docPath = fullName;
                        try { docPath = Path.GetFullPath(fullName); } catch { }
                        if (!string.Equals(fullName, filePath, StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(docPath, expectedPath, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (!doc.Saved)
                            doc.Save();
                        doc.Activate();
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
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[WordVariantScoring] TryActivateDocument: {ex.Message}");
            }

            return false;
        }

        private static void CloseOpenDocuments(WordApp wordApp)
        {
            if (wordApp == null)
                return;

            try
            {
                try { wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone; } catch { }
                int guard = 40;
                while (guard-- > 0)
                {
                    int count;
                    try { count = wordApp.Documents.Count; }
                    catch { return; }
                    if (count <= 0)
                        return;

                    WordDoc doc = null;
                    try
                    {
                        doc = wordApp.Documents[1];
                        doc.Close(SaveChanges: false);
                    }
                    catch
                    {
                        break;
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
                Debug.WriteLine($"[WordVariantScoring] CloseOpenDocuments: {ex.Message}");
            }
        }

        private static void ReportProgress(Action<string, int, int> progress, string message, int completed, int total)
        {
            if (progress == null)
                return;
            try { progress(message, completed, total); }
            catch (Exception ex)
            {
                Debug.WriteLine($"[WordVariantScoring] Progress: {ex.Message}");
            }
        }
    }
}
