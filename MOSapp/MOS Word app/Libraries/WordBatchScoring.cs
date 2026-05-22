using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Newtonsoft.Json;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WordDoc = Microsoft.Office.Interop.Word.Document;

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
                            doc.Close(SaveChanges: false);
                            break;
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
}
