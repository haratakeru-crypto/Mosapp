using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Libraries
{
    /// <summary>
    /// 1プロジェクト分のチェッカー呼び出し + ExcelScoreExplanation.Apply。
    /// 一括採点（ReviewPageWindow）の経路を正とし、開閉・flush・UI は呼び出し側に残す。
    /// </summary>
    public static class ExcelScoreProject
    {
        /// <summary>
        /// 指定ライブラリの全タスクを採点する。
        /// <paramref name="expectedFilePath"/> があるときは Impl / filePath 版を優先（一括）。
        /// null/空のときは公開メソッド（ActiveWorkbook）へフォールバック（Phase 2 その場用）。
        /// </summary>
        /// <param name="ensureExpectedWorkbookActive">
        /// 各タスク直前に呼ぶ任意コールバック（例: 対象ブックの Activate）。開閉自体は呼び出し側。
        /// </param>
        public static List<bool> ScoreProject(
            string libraryName,
            int taskCount,
            string expectedFilePath,
            int slotProjectId,
            out List<string> failReasons,
            Action<string> ensureExpectedWorkbookActive = null)
        {
            var results = new List<bool>();
            failReasons = new List<string>();

            try
            {
                Debug.WriteLine(
                    $"[ExcelScoreProject] ScoreProject libraryName={libraryName}, taskCount={taskCount}, slotProjectId={slotProjectId}");

                var parts = libraryName.Replace("ExcelChecker", "").Split('_');
                if (parts.Length < 2)
                {
                    Debug.WriteLine($"[ExcelScoreProject] Invalid library name format: {libraryName}");
                    FillAllUnavailable(results, failReasons, taskCount);
                    return results;
                }

                string groupId = parts[0];
                string projectId = parts[1];
                string namespaceName = $"Libraries.Group{groupId}";
                string fullTypeName = $"{namespaceName}.{libraryName}";

                Debug.WriteLine($"[ExcelScoreProject] Looking for type: {fullTypeName}");

                Type checkerType = ResolveCheckerType(libraryName, groupId, fullTypeName);
                if (checkerType == null)
                {
                    Debug.WriteLine($"[ExcelScoreProject] Checker type not found: {fullTypeName}");
                    FillAllUnavailable(results, failReasons, taskCount);
                    return results;
                }

                Debug.WriteLine($"[ExcelScoreProject] Type found: {checkerType.Name}");

#if DEBUG
                Debug.WriteLine($"[ExcelScoreProject] Available methods in {checkerType.Name}:");
                foreach (var method in checkerType.GetMethods())
                {
                    if (method.Name.StartsWith("CheckTask", StringComparison.Ordinal))
                        Debug.WriteLine($"[ExcelScoreProject] - {method.Name}");
                }
#endif

                object checkerInstance = Activator.CreateInstance(checkerType);

                var taskBindings = new CheckTaskMethodBinding[taskCount];
                for (int ti = 1; ti <= taskCount; ti++)
                    taskBindings[ti - 1] = CheckTaskMethodBinding.TryResolve(checkerType, groupId, projectId, ti);

                for (int i = 1; i <= taskCount; i++)
                {
                    if (!string.IsNullOrEmpty(expectedFilePath))
                        ensureExpectedWorkbookActive?.Invoke(expectedFilePath);

                    CheckTaskMethodBinding binding = taskBindings[i - 1];
                    if (binding == null)
                    {
                        Debug.WriteLine($"[ExcelScoreProject] No method found for task {i}, returning false");
                        results.Add(false);
                        failReasons.Add(ExcelScoreExplanation.UnavailableText);
                        continue;
                    }

                    string resolvedName = binding.ResolvedMethodName;

                    try
                    {
                        var taskSw = Stopwatch.StartNew();
                        ExcelScoreExplanation.ClearCheckerReason();
                        if (!binding.TryInvoke(checkerInstance, expectedFilePath, out bool invokeResult))
                        {
                            ExcelGradingPerf.Log(
                                "GradeTask.CheckerInvoke",
                                taskSw.ElapsedMilliseconds,
                                $"P{slotProjectId} T{i} invoke-failed");
                            ExcelScoreExplanation.ClearCheckerReason();
                            results.Add(false);
                            failReasons.Add(ExcelScoreExplanation.UnavailableText);
                            continue;
                        }

                        string checkerReason = ExcelScoreExplanation.TakeCheckerReason();
                        ExcelGradingPerf.Log(
                            "GradeTask.CheckerInvoke",
                            taskSw.ElapsedMilliseconds,
                            $"P{slotProjectId} T{i} passed={invokeResult}");
                        var gateSw = Stopwatch.StartNew();
                        // ログ・免除設定は画面上のスロット番号（VSTO の [Task N-...]）に合わせる
                        invokeResult = ExcelScoreExplanation.Apply(slotProjectId, i, invokeResult, checkerReason, out string reason);
                        ExcelGradingPerf.Log(
                            "GradeTask.DestructiveValidation",
                            gateSw.ElapsedMilliseconds,
                            $"P{slotProjectId} T{i} passed={invokeResult}");
                        Debug.WriteLine($"[ExcelScoreProject] Method {resolvedName} result: {invokeResult}");
                        results.Add(invokeResult);
                        failReasons.Add(invokeResult ? "" : (reason ?? ExcelScoreExplanation.RequirementMissText));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[ExcelScoreProject] Error invoking method {resolvedName}: {ex.Message}");
                        Debug.WriteLine($"StackTrace: {ex.StackTrace}");
                        results.Add(false);
                        failReasons.Add(ExcelScoreExplanation.UnavailableText);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExcelScoreProject] Error in ScoreProject: {ex.Message}");
                Debug.WriteLine($"StackTrace: {ex.StackTrace}");

                while (results.Count < taskCount)
                {
                    results.Add(false);
                    failReasons.Add(ExcelScoreExplanation.UnavailableText);
                }
            }

            return results;
        }

        private static void FillAllUnavailable(List<bool> results, List<string> failReasons, int taskCount)
        {
            for (int i = 0; i < taskCount; i++)
            {
                results.Add(false);
                failReasons.Add(ExcelScoreExplanation.UnavailableText);
            }
        }

        private static Type ResolveCheckerType(string libraryName, string groupId, string fullTypeName)
        {
            Type checkerType = null;

            var mainAssembly = AppDomain.CurrentDomain.GetAssemblies()
                .FirstOrDefault(a => a.GetName().Name == "MOSExcelMogiApp");

            if (mainAssembly != null)
            {
                try
                {
                    checkerType = mainAssembly.GetType(fullTypeName);

                    if (checkerType == null)
                    {
                        foreach (var type in mainAssembly.GetTypes())
                        {
                            if (type.Name == libraryName)
                            {
                                checkerType = type;
                                break;
                            }
                        }
                    }
                }
                catch (ReflectionTypeLoadException ex)
                {
                    Debug.WriteLine($"[ExcelScoreProject] ReflectionTypeLoadException: {ex.Message}");
                    if (ex.LoaderExceptions != null)
                    {
                        foreach (var loaderEx in ex.LoaderExceptions)
                            Debug.WriteLine($"[ExcelScoreProject] Loader exception: {loaderEx?.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ExcelScoreProject] Error getting type from main assembly: {ex.Message}");
                }
            }

            if (checkerType == null)
            {
                foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
                {
                    if (assembly == mainAssembly) continue;

                    try
                    {
                        checkerType = assembly.GetType(fullTypeName);
                        if (checkerType != null) break;

                        foreach (var type in assembly.GetTypes())
                        {
                            if (type.Name == libraryName)
                            {
                                checkerType = type;
                                break;
                            }
                        }
                        if (checkerType != null) break;
                    }
                    catch { }
                }
            }

            if (checkerType != null)
                return checkerType;

            Debug.WriteLine("[ExcelScoreProject] Type not found in loaded assemblies, trying to load DLL");

            string[] dllPaths =
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin", "Debug", "Libraries", $"Group{groupId}", $"{libraryName}.dll"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", $"Group{groupId}", $"{libraryName}.dll"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", "bin", "Debug", "net48", $"{libraryName}.dll")
            };

            foreach (string path in dllPaths)
            {
                Debug.WriteLine($"[ExcelScoreProject] Trying DLL path: {path}");
                if (!File.Exists(path))
                    continue;

                try
                {
                    Debug.WriteLine($"[ExcelScoreProject] Loading DLL from: {path}");
                    Assembly assembly = Assembly.LoadFrom(path);
                    checkerType = assembly.GetTypes().FirstOrDefault(t => t.Name == libraryName);
                    Debug.WriteLine($"[ExcelScoreProject] Found type from DLL: {checkerType?.Name ?? "null"}");
                    if (checkerType != null)
                        return checkerType;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ExcelScoreProject] Error loading DLL: {ex.Message}");
                }
            }

            Debug.WriteLine("[ExcelScoreProject] DLL not found in any of the searched paths");
            return null;
        }

        /// <summary>1タスク分の CheckTask / CheckTask_Impl の解決結果。タスクループ内の GetMethod 繰り返しを避ける。</summary>
        private sealed class CheckTaskMethodBinding
        {
            public MethodInfo ImplMethod;
            /// <summary>private bool CheckTask_* (string filePath) など、パス引き版。</summary>
            public MethodInfo FilePathMethod;
            public MethodInfo PublicMethod;
            public string ResolvedMethodName;

            public static string[] GetMethodNameCandidates(string groupId, string projectId, int taskIndex)
            {
                if (groupId == "1")
                {
                    return new[]
                    {
                        $"CheckTask_1_{projectId}_{taskIndex:D2}",
                        $"CheckTask_1_{projectId}_0{taskIndex}",
                        $"CheckTask_{groupId}_{projectId}_{taskIndex:D2}",
                        $"CheckTask_{groupId}_{projectId}_0{taskIndex}"
                    };
                }

                return new[]
                {
                    $"CheckTask_{groupId}_{projectId}_{taskIndex:D2}",
                    $"CheckTask_{groupId}_{projectId}_0{taskIndex}",
                    $"CheckTask_1_{projectId}_{taskIndex:D2}",
                    $"CheckTask_1_{projectId}_0{taskIndex}"
                };
            }

            public static CheckTaskMethodBinding TryResolve(Type checkerType, string groupId, string projectId, int taskIndex)
            {
                string[] methodNames = GetMethodNameCandidates(groupId, projectId, taskIndex);

                foreach (string methodName in methodNames)
                {
                    MethodInfo implMethod = checkerType.GetMethod(
                        methodName + "_Impl",
                        BindingFlags.Instance | BindingFlags.NonPublic);
                    MethodInfo filePathMethod = checkerType.GetMethod(
                        methodName,
                        BindingFlags.Instance | BindingFlags.NonPublic,
                        null,
                        new[] { typeof(string) },
                        null);
                    MethodInfo method = checkerType.GetMethod(methodName);
                    if (implMethod == null && filePathMethod == null && method == null)
                        continue;

                    return new CheckTaskMethodBinding
                    {
                        ImplMethod = implMethod,
                        FilePathMethod = filePathMethod,
                        PublicMethod = method,
                        ResolvedMethodName = methodName
                    };
                }

                return null;
            }

            /// <summary>既存ロジックと同じ分岐でチェッカーを呼び出す。パスがあれば ActiveWorkbook ではなく指定ファイルを採点する。</summary>
            public bool TryInvoke(object checkerInstance, string expectedFilePath, out bool result)
            {
                result = false;
                if (ImplMethod != null && !string.IsNullOrEmpty(expectedFilePath))
                {
                    result = (bool)ImplMethod.Invoke(checkerInstance, new object[] { expectedFilePath });
                    return true;
                }

                if (FilePathMethod != null && !string.IsNullOrEmpty(expectedFilePath))
                {
                    result = (bool)FilePathMethod.Invoke(checkerInstance, new object[] { expectedFilePath });
                    return true;
                }

                if (PublicMethod != null)
                {
                    result = (bool)PublicMethod.Invoke(checkerInstance, null);
                    return true;
                }

                return false;
            }
        }
    }
}
