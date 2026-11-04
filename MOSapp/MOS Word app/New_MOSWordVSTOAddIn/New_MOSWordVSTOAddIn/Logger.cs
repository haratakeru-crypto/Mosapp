using System;
using System.Globalization;
using System.IO;
using System.Text;
using Libraries;

namespace New_MOSWordVSTOAddIn
{
    /// <summary>
    /// Word 操作ログ。採点用は <see cref="LogTaskEvidence"/> の Project/Task 付き行のみ。
    /// <see cref="LogCommand"/> はデバッグ用（旧2括弧形式）。採点ロジックでは無視される。
    /// </summary>
    public static class Logger
    {
        private static readonly object _lockObject = new object();
        private static readonly string _logFileName = "mos_word_log.txt";
        private static string _logFilePath;
        private static int _currentProjectId = -1;
        private static int _currentTaskId = -1;
        private static int _currentAttemptNo;
        private static int _currentVariantSetNo;

        private static string LogFilePath
        {
            get
            {
                if (string.IsNullOrEmpty(_logFilePath))
                    _logFilePath = Path.Combine(Path.GetTempPath(), _logFileName);
                return _logFilePath;
            }
        }

        /// <summary>デバッグ用。形式: [timestamp] [commandId] Executed（採点では使用しない）</summary>
        public static void LogCommand(string commandId)
        {
            if (string.IsNullOrWhiteSpace(commandId))
                return;

            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string logEntry = $"[{timestamp}] [{commandId}] Executed";

                    using (var fileStream = new FileStream(
                        LogFilePath,
                        FileMode.Append,
                        FileAccess.Write,
                        FileShare.ReadWrite | FileShare.Delete))
                    using (var writer = new StreamWriter(fileStream, Encoding.UTF8))
                    {
                        writer.WriteLine(logEntry);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
            }
        }

        public static void ClearLog()
        {
            try
            {
                lock (_lockObject)
                {
                    if (File.Exists(LogFilePath))
                        File.Delete(LogFilePath);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error clearing log: {ex.Message}");
            }
        }

        public static string GetLogFilePath()
        {
            return LogFilePath;
        }

        /// <summary>Word 内で VSTO が稼働中であることを %TEMP% に記録する。</summary>
        public static void WriteVstoHeartbeat()
        {
            try
            {
                lock (_lockObject)
                {
                    string path = Path.Combine(Path.GetTempPath(), "mos_word_vsto_heartbeat.txt");
                    File.WriteAllText(path, DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString(), Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] WriteVstoHeartbeat: {ex.Message}");
            }
        }

        public static void SetCurrentTaskContext(int projectId, int taskId, int attemptNo, int variantSetNo = 0)
        {
            lock (_lockObject)
            {
                _currentProjectId = projectId;
                _currentTaskId = taskId;
                _currentAttemptNo = attemptNo < 0 ? 0 : attemptNo;
                _currentVariantSetNo = variantSetNo < 0 ? 0 : variantSetNo;
            }
        }

        /// <summary>現在の試験タスク文脈（未設定なら false）。</summary>
        public static bool TryGetCurrentTaskContext(out int projectId, out int taskId)
        {
            return TryGetCurrentTaskContext(out projectId, out taskId, out _);
        }

        /// <summary>現在の試験タスク文脈（未設定なら false）。類題は variantSetNo &gt; 0。</summary>
        public static bool TryGetCurrentTaskContext(out int projectId, out int taskId, out int variantSetNo)
        {
            lock (_lockObject)
            {
                projectId = _currentProjectId;
                taskId = _currentTaskId;
                variantSetNo = _currentVariantSetNo;
                return projectId > 0 && taskId > 0;
            }
        }

        /// <summary>
        /// 汎用操作ログ。[Task P-T-A] [Op] Type Detail（TaskStart は試験アプリが記録）。
        /// </summary>
        public static void LogOperation(string operationType, string detail)
        {
            if (string.IsNullOrWhiteSpace(operationType))
                return;
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string taskPrefix = "";
                    var context = ResolveCurrentContext(_currentProjectId, _currentTaskId);
                    if (context.IsValid)
                        taskPrefix = context.FormatExplicitTaskPrefix() + " ";
                    string logEntry = $"[{timestamp}] {taskPrefix}[Op] {operationType} {detail}".TrimEnd();
                    AppendLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] LogOperation: {ex.Message}");
            }
        }

        private static WordTaskLogContext ResolveCurrentContext(int fallbackProjectId, int fallbackTaskId)
        {
            bool hasLogger = _currentProjectId > 0 && _currentTaskId > 0;
            if (hasLogger && _currentVariantSetNo > 0)
            {
                return WordTaskLogContext.Create(
                    _currentProjectId, _currentTaskId, _currentAttemptNo, _currentVariantSetNo);
            }

            bool hasFile = TryReadCurrentTaskFile(out var fileContext) && fileContext.IsValid;
            if (hasFile && fileContext.IsVariant)
                return fileContext;

            if (hasLogger)
            {
                return WordTaskLogContext.Create(
                    _currentProjectId, _currentTaskId, _currentAttemptNo, 0);
            }

            if (hasFile)
                return fileContext;

            return WordTaskLogContext.Create(fallbackProjectId, fallbackTaskId, _currentAttemptNo, 0);
        }

        private static bool TryReadCurrentTaskFile(out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "mos_word_current_task.txt");
                if (!File.Exists(path))
                    return false;
                return WordTaskLogFormat.TryParseCurrentTaskLine(File.ReadAllText(path), out context)
                    && context.IsValid;
            }
            catch
            {
                return false;
            }
        }

        private static void AppendLine(string logEntry)
        {
            using (var fileStream = new FileStream(
                LogFilePath,
                FileMode.Append,
                FileAccess.Write,
                FileShare.ReadWrite | FileShare.Delete))
            using (var writer = new StreamWriter(fileStream, Encoding.UTF8))
            {
                writer.WriteLine(logEntry);
            }
        }

        /// <summary>
        /// 採点用ログ行を <c>mos_word_log.txt</c> に追記する。
        /// 教材: [timestamp] [ProjectN] [TaskN-M] [commandId] Executed
        /// 類題: [timestamp] [ProjectVn-N] [TaskVn-N-M] [commandId] Executed
        /// </summary>
        public static void LogTaskEvidence(int projectId, int taskId, string commandId)
        {
            if (string.IsNullOrWhiteSpace(commandId))
                return;

            try
            {
                lock (_lockObject)
                {
                    var current = ResolveCurrentContext(projectId, taskId);
                    var context = current.IsVariant
                        ? current
                        : WordTaskLogContext.Create(projectId, taskId, current.AttemptNo, 0);
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string logEntry = context.FormatScoringLine(timestamp, commandId);
                    AppendLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing scoring log: {ex.Message}");
            }
        }
    }
}

namespace Libraries
{
    /// <summary>
    /// 教材 / 類題を区別する操作ログのタスク文脈。
    /// </summary>
    public struct WordTaskLogContext : IEquatable<WordTaskLogContext>
    {
        public int ProjectId;
        public int TaskId;
        public int AttemptNo;
        public int VariantSetNo;
        public int ExemptFlags;

        public bool IsValid => ProjectId > 0 && TaskId > 0;
        public bool IsVariant => VariantSetNo > 0;

        public static WordTaskLogContext Create(
            int projectId,
            int taskId,
            int attemptNo = 0,
            int variantSetNo = 0,
            int exemptFlags = 0)
        {
            return new WordTaskLogContext
            {
                ProjectId = projectId,
                TaskId = taskId,
                AttemptNo = attemptNo < 0 ? 0 : attemptNo,
                VariantSetNo = variantSetNo < 0 ? 0 : variantSetNo,
                ExemptFlags = exemptFlags < 0 ? 0 : exemptFlags
            };
        }

        public string FormatTaskToken()
        {
            if (VariantSetNo > 0)
                return "V" + VariantSetNo.ToString(CultureInfo.InvariantCulture)
                    + "-" + ProjectId.ToString(CultureInfo.InvariantCulture)
                    + "-" + TaskId.ToString(CultureInfo.InvariantCulture)
                    + "-" + AttemptNo.ToString(CultureInfo.InvariantCulture);
            return ProjectId.ToString(CultureInfo.InvariantCulture)
                + "-" + TaskId.ToString(CultureInfo.InvariantCulture)
                + "-" + AttemptNo.ToString(CultureInfo.InvariantCulture);
        }

        public string FormatCurrentTaskLine()
        {
            string line = ProjectId.ToString(CultureInfo.InvariantCulture)
                + "," + TaskId.ToString(CultureInfo.InvariantCulture)
                + "," + ExemptFlags.ToString(CultureInfo.InvariantCulture)
                + "," + AttemptNo.ToString(CultureInfo.InvariantCulture);
            if (VariantSetNo > 0)
                line += ",V" + VariantSetNo.ToString(CultureInfo.InvariantCulture);
            return line;
        }

        public string FormatEvidenceProjectToken()
        {
            if (VariantSetNo > 0)
                return "ProjectV" + VariantSetNo.ToString(CultureInfo.InvariantCulture)
                    + "-" + ProjectId.ToString(CultureInfo.InvariantCulture);
            return "Project" + ProjectId.ToString(CultureInfo.InvariantCulture);
        }

        public string FormatEvidenceTaskToken()
        {
            if (VariantSetNo > 0)
                return "TaskV" + VariantSetNo.ToString(CultureInfo.InvariantCulture)
                    + "-" + ProjectId.ToString(CultureInfo.InvariantCulture)
                    + "-" + TaskId.ToString(CultureInfo.InvariantCulture);
            return "Task" + ProjectId.ToString(CultureInfo.InvariantCulture)
                + "-" + TaskId.ToString(CultureInfo.InvariantCulture);
        }

        public string FormatScoringLine(string timestamp, string commandId)
        {
            return "[" + timestamp + "] [" + FormatEvidenceProjectToken() + "] ["
                + FormatEvidenceTaskToken() + "] [" + commandId + "] Executed";
        }

        public string FormatDestructivePrefix()
        {
            if (VariantSetNo > 0)
                return "V" + VariantSetNo.ToString(CultureInfo.InvariantCulture)
                    + "," + ProjectId.ToString(CultureInfo.InvariantCulture)
                    + "," + TaskId.ToString(CultureInfo.InvariantCulture)
                    + "," + AttemptNo.ToString(CultureInfo.InvariantCulture)
                    + ":";
            return ProjectId.ToString(CultureInfo.InvariantCulture)
                + "," + TaskId.ToString(CultureInfo.InvariantCulture)
                + "," + AttemptNo.ToString(CultureInfo.InvariantCulture)
                + ":";
        }

        public string FormatExplicitTaskPrefix()
        {
            return "[Task " + FormatTaskToken() + "]";
        }

        public bool MatchesProjectScope(int projectId, int variantSetNo)
        {
            return ProjectId == projectId && VariantSetNo == (variantSetNo < 0 ? 0 : variantSetNo);
        }

        public bool Equals(WordTaskLogContext other)
        {
            return ProjectId == other.ProjectId
                && TaskId == other.TaskId
                && AttemptNo == other.AttemptNo
                && VariantSetNo == other.VariantSetNo
                && ExemptFlags == other.ExemptFlags;
        }

        public override bool Equals(object obj)
        {
            return obj is WordTaskLogContext other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = ProjectId;
                hash = (hash * 397) ^ TaskId;
                hash = (hash * 397) ^ AttemptNo;
                hash = (hash * 397) ^ VariantSetNo;
                hash = (hash * 397) ^ ExemptFlags;
                return hash;
            }
        }
    }

    public static class WordTaskLogFormat
    {
        public static bool TryParseTaskToken(string token, out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            if (string.IsNullOrWhiteSpace(token))
                return false;

            string[] parts = token.Trim().Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
                return false;

            int offset = 0;
            int variantSetNo = 0;
            if (parts[0].Length >= 2 && (parts[0][0] == 'V' || parts[0][0] == 'v'))
            {
                if (parts.Length < 3)
                    return false;
                if (!int.TryParse(parts[0].Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out variantSetNo)
                    || variantSetNo < 1)
                    return false;
                offset = 1;
            }

            if (!int.TryParse(parts[offset], NumberStyles.Integer, CultureInfo.InvariantCulture, out int projectId)
                || projectId <= 0)
                return false;
            if (!int.TryParse(parts[offset + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int taskId)
                || taskId <= 0)
                return false;

            int attemptNo = 0;
            if (parts.Length > offset + 2)
            {
                if (!int.TryParse(parts[offset + 2], NumberStyles.Integer, CultureInfo.InvariantCulture, out attemptNo)
                    || attemptNo < 0)
                    attemptNo = 0;
            }

            context = WordTaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo);
            return true;
        }

        public static bool TryParseCurrentTaskLine(string line, out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            if (string.IsNullOrWhiteSpace(line))
                return false;

            string[] parts = line.Trim().Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
                return false;

            if (!int.TryParse(parts[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int projectId)
                || projectId <= 0)
                return false;
            if (!int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int taskId)
                || taskId <= 0)
                return false;

            int exemptFlags = 0;
            if (parts.Length >= 3)
                int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out exemptFlags);

            int attemptNo = 0;
            if (parts.Length >= 4)
                int.TryParse(parts[3].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out attemptNo);

            int variantSetNo = 0;
            if (parts.Length >= 5)
            {
                string marker = parts[4].Trim();
                if (marker.Length < 2 || (marker[0] != 'V' && marker[0] != 'v'))
                    return false;
                if (!int.TryParse(marker.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out variantSetNo)
                    || variantSetNo < 1)
                    return false;
            }

            context = WordTaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo, exemptFlags);
            return true;
        }

        public static bool TryParseTaskStartLine(string line, out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            if (string.IsNullOrWhiteSpace(line))
                return false;

            int idx = line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
                return false;

            string token = line.Substring(idx + 11).Trim();
            int space = token.IndexOf(' ');
            if (space > 0)
                token = token.Substring(0, space);
            return TryParseTaskToken(token, out context);
        }

        public static bool TryParseExplicitTaskPrefix(string line, out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            if (string.IsNullOrWhiteSpace(line))
                return false;

            int idx = line.IndexOf("[Task ", StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
                return false;

            int start = idx + 6;
            int end = line.IndexOf(']', start);
            if (end <= start)
                return false;

            return TryParseTaskToken(line.Substring(start, end - start), out context);
        }

        public static bool TryParseDestructiveLinePrefix(string line, out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            if (string.IsNullOrWhiteSpace(line))
                return false;

            int colon = line.IndexOf(':');
            if (colon <= 0)
                return false;

            string key = line.Substring(0, colon);
            string[] parts = key.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3)
                return false;

            int offset = 0;
            int variantSetNo = 0;
            string first = parts[0].Trim();
            if (first.Length >= 2 && (first[0] == 'V' || first[0] == 'v'))
            {
                if (parts.Length < 4)
                    return false;
                if (!int.TryParse(first.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out variantSetNo)
                    || variantSetNo < 1)
                    return false;
                offset = 1;
            }

            if (!int.TryParse(parts[offset].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int projectId)
                || projectId <= 0)
                return false;
            if (!int.TryParse(parts[offset + 1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int taskId)
                || taskId <= 0)
                return false;

            int attemptNo = 0;
            if (parts.Length > offset + 2)
            {
                if (!int.TryParse(parts[offset + 2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out attemptNo)
                    || attemptNo < 0)
                    attemptNo = 0;
            }

            context = WordTaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo);
            return true;
        }
    }
}
