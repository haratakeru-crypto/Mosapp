using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Libraries
{
    /// <summary>
    /// VSTO が追記する <c>mos_word_log.txt</c> を読み、採点用 Project/Task 付き行のみを解釈する。
    /// </summary>
    public static class LogReader
    {
        private const string LegacyTaskEvidenceFileName = "mos_word_task_evidence.txt";

        /// <summary>
        /// 採点用1行: [timestamp] [ProjectN] [TaskN-M] [CommandId] Executed。Task の N は Project と一致すること。
        /// </summary>
        private static readonly Regex ScoringLineRegex = new Regex(
            @"^\[(?<ts>[^\]]+)\]\s+\[Project(?<proj>\d+)\]\s+\[Task(?<tproj>\d+)-(?<task>\d+)\]\s+\[(?<cmd>[^\]]+)\]\s+Executed\s*$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static string GetLogFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_word_log.txt");
        }

        /// <summary>互換: 旧証跡ファイル名。移行後は使用しない。</summary>
        [Obsolete("採点ログは GetLogFilePath() のみを使用してください。")]
        public static string GetTaskEvidenceLogPath()
        {
            return Path.Combine(Path.GetTempPath(), LegacyTaskEvidenceFileName);
        }

        public static string GetTaskEvidenceMarker(int projectId, int taskId)
        {
            return $"[Task{projectId}-{taskId}]";
        }

        public static void ClearLog()
        {
            try
            {
                string path = GetLogFilePath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] Error clearing log: {ex.Message}");
            }
        }

        /// <summary>旧 API 名。全リセット時は <see cref="ClearLog"/> のみで足りる。</summary>
        public static void ClearTaskEvidence()
        {
            try
            {
                string legacy = Path.Combine(Path.GetTempPath(), LegacyTaskEvidenceFileName);
                if (File.Exists(legacy))
                    File.Delete(legacy);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] ClearTaskEvidence (legacy file): {ex.Message}");
            }
        }

        /// <summary>個別リセット: <c>[ProjectN]</c> の採点行のみ削除（他行は維持）。</summary>
        public static void ClearTaskEvidenceForProject(int projectId)
        {
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return;

            try
            {
                var lines = File.ReadAllLines(path, Encoding.UTF8);
                var kept = lines.Where(l =>
                {
                    if (string.IsNullOrWhiteSpace(l))
                        return true;
                    if (!TryParseScoringLine(l, out var e) || !e.IsValid)
                        return true;
                    return e.ProjectId != projectId;
                }).ToArray();

                if (kept.Length == lines.Length)
                    return;

                if (kept.Length == 0 || kept.All(string.IsNullOrWhiteSpace))
                {
                    File.Delete(path);
                    return;
                }

                File.WriteAllLines(path, kept, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] ClearTaskEvidenceForProject: {ex.Message}");
            }
        }

        /// <summary>テスト・デバッグ用。採点と同一形式で追記。</summary>
        public static void AppendTaskEvidence(int projectId, int taskId, string commandId)
        {
            if (string.IsNullOrWhiteSpace(commandId))
                return;

            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string line =
                    $"[{timestamp}] [Project{projectId}] [Task{projectId}-{taskId}] [{commandId}] Executed";
                File.AppendAllText(GetLogFilePath(), line + Environment.NewLine, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] AppendTaskEvidence: {ex.Message}");
            }
        }

        public static bool HasTaskEvidence(int projectId, int taskId, string commandId)
        {
            return CountTaskEvidence(projectId, taskId, commandId) > 0;
        }

        public static bool HasTaskEvidenceAtLeast(int projectId, int taskId, string commandId, int minCount)
        {
            return CountTaskEvidence(projectId, taskId, commandId) >= minCount;
        }

        public static int CountTaskEvidenceForProject(int projectId, string commandId)
        {
            if (string.IsNullOrWhiteSpace(commandId))
                return 0;

            string path = GetLogFilePath();
            if (!File.Exists(path))
                return 0;

            int count = 0;
            try
            {
                foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;
                    if (!TryParseScoringLine(line, out var e) || !e.IsValid)
                        continue;
                    if (e.ProjectId != projectId)
                        continue;
                    if (string.Equals(e.CommandId, commandId, StringComparison.OrdinalIgnoreCase))
                        count++;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] CountTaskEvidenceForProject: {ex.Message}");
            }

            return count;
        }

        public static bool HasTaskEvidenceForProjectAtLeast(int projectId, string commandId, int minCount)
        {
            return CountTaskEvidenceForProject(projectId, commandId) >= minCount;
        }

        private static int CountTaskEvidence(int projectId, int taskId, string commandId)
        {
            if (string.IsNullOrWhiteSpace(commandId))
                return 0;

            string path = GetLogFilePath();
            if (!File.Exists(path))
                return 0;

            int count = 0;
            try
            {
                foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;
                    if (!TryParseScoringLine(line, out var e) || !e.IsValid)
                        continue;
                    if (e.ProjectId != projectId || e.TaskId != taskId)
                        continue;
                    if (string.Equals(e.CommandId, commandId, StringComparison.OrdinalIgnoreCase))
                        count++;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] CountTaskEvidence: {ex.Message}");
            }

            return count;
        }

        public static List<LogEntry> ReadAllLogEntries()
        {
            var entries = new List<LogEntry>();
            string logFilePath = GetLogFilePath();
            if (!File.Exists(logFilePath))
                return entries;

            try
            {
                foreach (string line in File.ReadAllLines(logFilePath, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;
                    if (TryParseScoringLine(line, out var e) && e.IsValid)
                    {
                        entries.Add(new LogEntry
                        {
                            Timestamp = e.Timestamp,
                            CommandId = e.CommandId,
                            ProjectId = e.ProjectId,
                            TaskId = e.TaskId,
                            IsScoringLine = true
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] Error reading log file: {ex.Message}");
            }

            return entries;
        }

        public static bool HasAnyTaskEvidence(int projectId, int taskId, params string[] commandIds)
        {
            if (commandIds == null || commandIds.Length == 0)
                return false;
            foreach (string id in commandIds)
            {
                if (string.IsNullOrWhiteSpace(id))
                    continue;
                if (HasTaskEvidence(projectId, taskId, id.Trim()))
                    return true;
            }
            return false;
        }

        private static bool TryParseScoringLine(string line, out ScoringLogFields fields)
        {
            fields = default;
            if (string.IsNullOrWhiteSpace(line))
                return false;

            Match m = ScoringLineRegex.Match(line);
            if (!m.Success)
                return false;

            if (!DateTime.TryParse(m.Groups["ts"].Value, out DateTime ts))
                return false;

            if (!int.TryParse(m.Groups["proj"].Value, out int proj))
                return false;
            if (!int.TryParse(m.Groups["tproj"].Value, out int tproj))
                return false;
            if (!int.TryParse(m.Groups["task"].Value, out int taskNum))
                return false;

            string cmd = m.Groups["cmd"].Value;
            if (string.IsNullOrWhiteSpace(cmd))
                return false;

            bool coherent = proj == tproj;
            fields = new ScoringLogFields
            {
                Timestamp = ts,
                ProjectId = proj,
                TaskId = taskNum,
                CommandId = cmd,
                IsValid = coherent
            };
            return true;
        }

        private struct ScoringLogFields
        {
            public DateTime Timestamp;
            public int ProjectId;
            public int TaskId;
            public string CommandId;
            public bool IsValid;
        }

        public class LogEntry
        {
            public DateTime Timestamp { get; set; }
            public string CommandId { get; set; }
            public int ProjectId { get; set; }
            public int TaskId { get; set; }
            public bool IsScoringLine { get; set; }

            public override string ToString()
            {
                if (IsScoringLine)
                    return $"[{Timestamp:yyyy-MM-dd HH:mm:ss}] [Project{ProjectId}] [Task{ProjectId}-{TaskId}] [{CommandId}] Executed";
                return $"[{Timestamp:yyyy-MM-dd HH:mm:ss}] [{CommandId}] Executed";
            }
        }
    }
}
