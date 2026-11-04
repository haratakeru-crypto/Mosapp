using System;
using System.IO;
using System.Text;

namespace ExcelAddIn1
{
    /// <summary>
    /// Excel 操作ログを %TEMP%\mos_excel_log.txt に記録する。
    /// [Op] / [TaskStart] は採点用。Executed は Debug、または %TEMP%\mos_excel_diag.on があるときだけ書く。
    /// </summary>
    public static class Logger
    {
        private static readonly object LockObject = new object();
        private static readonly string LogFilePath = Path.Combine(Path.GetTempPath(), "mos_excel_log.txt");
        private static int _currentProjectId = -1;
        private static int _currentTaskId = -1;
        private static int _currentAttemptNo = 1;
        private static int _currentVariantSetNo = 0;

        public static void LogCommand(string commandId)
        {
            if (!ExcelDiagFlags.IsVerbose)
                return;

            try
            {
                lock (LockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    AppendToFile($"[{timestamp}] [{commandId}] Executed");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelLogger] Error writing log: " + ex.Message);
            }
        }

        public static void SetCurrentTaskContext(int projectId, int taskId, int attemptNo)
        {
            SetCurrentTaskContext(projectId, taskId, attemptNo, _currentVariantSetNo);
        }

        public static void SetCurrentTaskContext(int projectId, int taskId, int attemptNo, int variantSetNo)
        {
            lock (LockObject)
            {
                _currentProjectId = projectId;
                _currentTaskId = taskId;
                _currentAttemptNo = attemptNo < 1 ? 1 : attemptNo;
                _currentVariantSetNo = variantSetNo < 0 ? 0 : variantSetNo;
            }
        }

        /// <summary>現在のタスク文脈（<see cref="LogOperation"/> の接頭辞に使われる値）。</summary>
        public static void GetCurrentTaskContext(out int projectId, out int taskId, out int attemptNo)
        {
            GetCurrentTaskContext(out projectId, out taskId, out attemptNo, out _);
        }

        public static void GetCurrentTaskContext(out int projectId, out int taskId, out int attemptNo, out int variantSetNo)
        {
            lock (LockObject)
            {
                projectId = _currentProjectId;
                taskId = _currentTaskId;
                attemptNo = _currentAttemptNo;
                variantSetNo = _currentVariantSetNo;
            }
        }

        /// <summary>
        /// 一時的にタスク文脈を差し替えて <paramref name="action"/> を実行する（同一スレッド・再入可能）。
        /// Backstage 開閉など、ログ記録時だけ過去のタスク文脈を使う用途。
        /// </summary>
        public static void RunWithTaskContext(int projectId, int taskId, int attemptNo, Action action)
        {
            RunWithTaskContext(projectId, taskId, attemptNo, -1, action);
        }

        public static void RunWithTaskContext(int projectId, int taskId, int attemptNo, int variantSetNo, Action action)
        {
            if (action == null) return;
            lock (LockObject)
            {
                int prevP = _currentProjectId;
                int prevT = _currentTaskId;
                int prevA = _currentAttemptNo;
                int prevV = _currentVariantSetNo;
                try
                {
                    if (projectId > 0 && taskId > 0)
                    {
                        _currentProjectId = projectId;
                        _currentTaskId = taskId;
                        _currentAttemptNo = attemptNo < 1 ? 1 : attemptNo;
                        if (variantSetNo >= 0)
                            _currentVariantSetNo = variantSetNo;
                    }
                    action();
                }
                finally
                {
                    _currentProjectId = prevP;
                    _currentTaskId = prevT;
                    _currentAttemptNo = prevA;
                    _currentVariantSetNo = prevV;
                }
            }
        }

        public static void LogTaskStart(int projectId, int taskId, int attemptNo = 1)
        {
            LogTaskStart(projectId, taskId, attemptNo, _currentVariantSetNo);
        }

        public static void LogTaskStart(int projectId, int taskId, int attemptNo, int variantSetNo)
        {
            try
            {
                lock (LockObject)
                {
                    if (attemptNo < 1) attemptNo = 1;
                    if (variantSetNo < 0) variantSetNo = 0;
                    _currentProjectId = projectId;
                    _currentTaskId = taskId;
                    _currentAttemptNo = attemptNo;
                    _currentVariantSetNo = variantSetNo;
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    var context = TaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo);
                    AppendToFile($"[{timestamp}] [TaskStart] {context.FormatTaskToken()}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelLogger] Error writing task start: " + ex.Message);
            }
        }

        public static void LogOperation(string operationType, string detail)
        {
            try
            {
                lock (LockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string taskPrefix = "";
                    if (_currentProjectId > 0 && _currentTaskId > 0)
                    {
                        var context = TaskLogContext.Create(_currentProjectId, _currentTaskId, _currentAttemptNo, _currentVariantSetNo);
                        taskPrefix = context.FormatExplicitTaskPrefix() + " ";
                    }
                    AppendToFile($"[{timestamp}] {taskPrefix}[Op] {operationType} {detail}".TrimEnd());
                }

                // 単語帳: gallery/menu は Ribbon フック不可のため、操作ログからキーを補完
                try
                {
                    if (string.Equals(operationType, "AddConditionalFormat", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(operationType, "SetConditionalFormat", StringComparison.OrdinalIgnoreCase))
                        VocabLogger.LogKey("ConditionalFormattingMenu");
                    else if (string.Equals(operationType, "SetPageMargins", StringComparison.OrdinalIgnoreCase))
                        VocabLogger.LogKey("PageMarginsGallery");
                    else if (string.Equals(operationType, "ManageNamedRange", StringComparison.OrdinalIgnoreCase))
                        VocabLogger.LogKey("NameUseInFormula");
                    else if (string.Equals(operationType, "SetFreezePanes", StringComparison.OrdinalIgnoreCase))
                        VocabLogger.LogKey("FreezePanes");
                    else if (string.Equals(operationType, "SortOrFilter", StringComparison.OrdinalIgnoreCase))
                    {
                        VocabLogger.LogKey("SortDialog");
                        VocabLogger.LogKey("SortAscendingExcel");
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelLogger] Error writing operation log: " + ex.Message);
            }
        }

        public static void ClearLog()
        {
            try
            {
                lock (LockObject)
                {
                    if (File.Exists(LogFilePath))
                    {
                        File.Delete(LogFilePath);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelLogger] Error clearing log: " + ex.Message);
            }
        }

        public static string GetLogFilePath()
        {
            return LogFilePath;
        }

        private static void AppendToFile(string line)
        {
            using (var stream = new FileStream(
                LogFilePath,
                FileMode.Append,
                FileAccess.Write,
                FileShare.ReadWrite | FileShare.Delete))
            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                writer.WriteLine(line);
            }
        }
    }
}
