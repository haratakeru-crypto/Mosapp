using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Libraries
{
    /// <summary>
    /// VSTO が追記する <c>mos_word_log.txt</c> を読み、採点用 Project/Task 付き行のみを解釈する。
    /// </summary>
    public sealed class VstoEvidenceFlushAck
    {
        public bool Completed { get; set; }
        public long RequestId { get; set; }
        public int ProjectId { get; set; }
        public int TaskId { get; set; }
        public int AttemptNo { get; set; }
        public bool EvidenceWritten { get; set; }
    }

    public static class LogReader
    {
        private const string LegacyTaskEvidenceFileName = "mos_word_task_evidence.txt";
        private const string EvidenceFlushFileName = "mos_word_flush_evidence.txt";
        private const string EvidenceFlushAckFileName = "mos_word_flush_evidence.ack.txt";
        private const string CloseNavigationFileName = "mos_word_close_navigation.txt";
        private const string BatchCacheEnabledKey = "MOS.Word.LogReader.BatchCacheEnabled";
        private const string BatchCacheGenerationKey = "MOS.Word.LogReader.BatchCacheGeneration";

        private static readonly object CacheSync = new object();
        private static CachedLines LogFileCache;
        private static CachedLines DestructiveFileCache;

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

        /// <summary>VSTO アドインが Word 内で動作中であることを示すハートビートファイル。</summary>
        public static string GetVstoHeartbeatPath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_word_vsto_heartbeat.txt");
        }

        /// <summary>直近で VSTO がハートビートを更新していれば true（既定 5 分以内）。</summary>
        public static bool IsVstoHeartbeatFresh(int maxAgeSeconds = 300)
        {
            try
            {
                string path = GetVstoHeartbeatPath();
                if (!File.Exists(path))
                    return false;

                string text = File.ReadAllText(path, Encoding.UTF8).Trim();
                if (!long.TryParse(text, out long unixMs))
                    return false;

                double ageSec = (DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() - unixMs) / 1000.0;
                return ageSec >= 0 && ageSec <= maxAgeSeconds;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] IsVstoHeartbeatFresh: {ex.Message}");
                return false;
            }
        }

        /// <summary>Word 再起動前に古い心拍を消し、前セッションの誤検知を防ぐ。</summary>
        public static void ClearVstoHeartbeat()
        {
            try
            {
                string path = GetVstoHeartbeatPath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] ClearVstoHeartbeat: {ex.Message}");
            }
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

        public static void BeginBatchScoringLogCache()
        {
            AppDomain.CurrentDomain.SetData(BatchCacheEnabledKey, true);
            InvalidateBatchScoringLogCache();
        }

        public static void EndBatchScoringLogCache()
        {
            AppDomain.CurrentDomain.SetData(BatchCacheEnabledKey, false);
            InvalidateBatchScoringLogCache();
            lock (CacheSync)
            {
                LogFileCache = default;
                DestructiveFileCache = default;
            }
        }

        public static void InvalidateBatchScoringLogCache()
        {
            object current = AppDomain.CurrentDomain.GetData(BatchCacheGenerationKey);
            int generation = current is int value ? value : 0;
            AppDomain.CurrentDomain.SetData(BatchCacheGenerationKey, unchecked(generation + 1));
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
            finally
            {
                InvalidateBatchScoringLogCache();
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

        /// <summary>
        /// 個別リセット: 対象プロジェクトの証跡（Executed）・TaskStart・[Op] を削除し、他プロジェクトの行は維持する。
        /// </summary>
        public static void ClearTaskEvidenceForProject(int projectId)
        {
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return;

            try
            {
                var lines = File.ReadAllLines(path, Encoding.UTF8);
                var kept = lines.Where(l => !IsLogLineOwnedByProject(l, projectId)).ToArray();

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
            finally
            {
                InvalidateBatchScoringLogCache();
            }
        }

        /// <summary>個別リセット: 対象プロジェクトの破壊検知エラー行のみ削除する。</summary>
        public static void ClearDestructiveLogForProject(int projectId)
        {
            string path = GetDestructiveErrorLogPath();
            if (!File.Exists(path))
                return;

            string prefix = projectId.ToString() + ",";
            try
            {
                var lines = File.ReadAllLines(path, Encoding.UTF8);
                var kept = lines.Where(l =>
                    string.IsNullOrWhiteSpace(l)
                    || !l.StartsWith(prefix, StringComparison.Ordinal)).ToArray();

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
                System.Diagnostics.Debug.WriteLine($"[LogReader] ClearDestructiveLogForProject: {ex.Message}");
            }
            finally
            {
                InvalidateBatchScoringLogCache();
            }
        }

        private static bool IsLogLineOwnedByProject(string line, int projectId)
        {
            if (string.IsNullOrWhiteSpace(line))
                return false;

            if (TryParseScoringLine(line, out var e) && e.IsValid && e.ProjectId == projectId)
                return true;

            if (line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                ParseTaskStart(line, out int p, out _, out _);
                return p == projectId;
            }

            string taskMarker = "[Task " + projectId + "-";
            if (line.IndexOf(taskMarker, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            return false;
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
            finally
            {
                InvalidateBatchScoringLogCache();
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
                foreach (string line in ReadCachedLines(path))
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
                foreach (string line in ReadCachedLines(path))
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
                foreach (string line in ReadCachedLines(logFilePath))
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

        // --- 破壊的操作検知（TaskStart / Op / destructive / snapshot）---

        public static string GetCurrentTaskFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_word_current_task.txt");
        }

        public static string GetDestructiveErrorLogPath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_word_destructive_errors.log");
        }

        public static string GetSnapshotFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_word_snapshot.txt");
        }

        public static string GetEvidenceFlushFilePath()
        {
            return Path.Combine(Path.GetTempPath(), EvidenceFlushFileName);
        }

        public static string GetEvidenceFlushAckFilePath()
        {
            return Path.Combine(Path.GetTempPath(), EvidenceFlushAckFileName);
        }

        public static string GetCloseNavigationFilePath()
        {
            return Path.Combine(Path.GetTempPath(), CloseNavigationFileName);
        }

        /// <summary>
        /// VSTO にナビゲーションウィンドウを閉じるよう依頼する（開いているときのみ閉じる）。処理完了まで待機する。
        /// </summary>
        public static void RequestCloseNavigationPaneIfOpen(int timeoutMs = 600)
        {
            string path = GetCloseNavigationFilePath();
            try
            {
                File.WriteAllText(path, DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString(), Encoding.UTF8);
                var sw = Stopwatch.StartNew();
                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    if (!File.Exists(path))
                        return;
                    Thread.Sleep(30);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] RequestCloseNavigationPaneIfOpen: {ex.Message}");
            }
        }

        /// <summary>
        /// 採点直前に VSTO へ ShowAll 等のポーリング同期を依頼し、処理完了まで待機する。
        /// </summary>
        public static void RequestVstoEvidenceFlush(int timeoutMs = 1200)
        {
            TryRequestVstoEvidenceFlush(0, 0, 0, timeoutMs);
        }

        /// <summary>
        /// VSTO 証跡確定を依頼する。一致する応答があれば true。
        /// </summary>
        public static bool TryRequestVstoEvidenceFlush(int timeoutMs = 1200)
        {
            return TryRequestVstoEvidenceFlush(0, 0, 0, timeoutMs).Completed;
        }

        /// <summary>
        /// 指定タスクの VSTO 証跡確定を依頼する。ack の requestId とタスクが一致したときだけ完了とする。
        /// </summary>
        public static VstoEvidenceFlushAck TryRequestVstoEvidenceFlush(int projectId, int taskId, int attemptNo, int timeoutMs = 1200)
        {
            long requestId = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            string path = GetEvidenceFlushFilePath();
            string ackPath = GetEvidenceFlushAckFilePath();
            var result = new VstoEvidenceFlushAck
            {
                RequestId = requestId,
                ProjectId = projectId,
                TaskId = taskId,
                AttemptNo = attemptNo
            };
            try
            {
                try
                {
                    if (File.Exists(ackPath))
                        File.Delete(ackPath);
                }
                catch { }

                string payload = $"{requestId},{projectId},{taskId},{attemptNo},{requestId}";
                File.WriteAllText(path, payload, new UTF8Encoding(false));
                var sw = Stopwatch.StartNew();
                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    if (TryReadEvidenceFlushAck(ackPath, requestId, projectId, taskId, attemptNo, out bool evidenceWritten))
                    {
                        result.Completed = true;
                        result.EvidenceWritten = evidenceWritten;
                        break;
                    }
                    Thread.Sleep(30);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] RequestVstoEvidenceFlush: {ex.Message}");
            }
            finally
            {
                try
                {
                    if (File.Exists(ackPath))
                        File.Delete(ackPath);
                }
                catch { }
                InvalidateBatchScoringLogCache();
            }
            return result;
        }

        private static bool TryReadEvidenceFlushAck(
            string ackPath,
            long requestId,
            int projectId,
            int taskId,
            int attemptNo,
            out bool evidenceWritten)
        {
            evidenceWritten = false;
            try
            {
                if (string.IsNullOrEmpty(ackPath) || !File.Exists(ackPath))
                    return false;
                string[] parts = File.ReadAllText(ackPath, Encoding.UTF8).Trim()
                    .Split(new[] { ',' }, StringSplitOptions.None);
                if (parts.Length < 6)
                    return false;
                if (!long.TryParse(parts[0].Trim(), out long ackRequestId) || ackRequestId != requestId)
                    return false;
                if (!int.TryParse(parts[1].Trim(), out int ackProjectId) || ackProjectId != projectId)
                    return false;
                if (!int.TryParse(parts[2].Trim(), out int ackTaskId) || ackTaskId != taskId)
                    return false;
                if (!int.TryParse(parts[3].Trim(), out int ackAttemptNo) || ackAttemptNo != attemptNo)
                    return false;
                if (!int.TryParse(parts[4].Trim(), out int processed) || processed != 1)
                    return false;
                evidenceWritten = int.TryParse(parts[5].Trim(), out int written) && written == 1;
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] TryReadEvidenceFlushAck: {ex.Message}");
                return false;
            }
        }

        public static void ClearCurrentTaskFile()
        {
            try
            {
                string path = GetCurrentTaskFilePath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] ClearCurrentTaskFile: {ex.Message}");
            }
        }

        public static void ClearDestructiveLog()
        {
            try
            {
                string path = GetDestructiveErrorLogPath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] ClearDestructiveLog: {ex.Message}");
            }
            finally
            {
                InvalidateBatchScoringLogCache();
            }
        }

        public static void ClearSnapshot()
        {
            try
            {
                string path = GetSnapshotFilePath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] ClearSnapshot: {ex.Message}");
            }
        }

        /// <summary>アプリ側がタスク表示時に記録（記録主体は試験アプリ）。</summary>
        public static void LogTaskStart(int projectId, int taskId, int attemptNo)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string line = $"[{timestamp}] [TaskStart] {projectId}-{taskId}-{attemptNo}";
                File.AppendAllText(GetLogFilePath(), line + Environment.NewLine, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] LogTaskStart: {ex.Message}");
            }
            finally
            {
                InvalidateBatchScoringLogCache();
            }
        }

        public static List<string> GetOperationsForTask(int projectId, int taskId, int attemptNo)
        {
            var result = new List<string>();
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return result;

            try
            {
                string[] lines = ReadCachedLines(path);
                int curP = -1, curT = -1, curA = 0;
                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;
                    if (line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ParseTaskStart(line, out curP, out curT, out curA);
                        continue;
                    }
                    int opIdx = line.IndexOf("[Op]", StringComparison.OrdinalIgnoreCase);
                    if (opIdx < 0)
                        continue;
                    // 行に [Task P-T-A] があれば VSTO 記録時のタスクを優先（TaskStart より後に並んでも誤帰属しない）
                    if (TryParseExplicitOpTask(line, opIdx, out int opP, out int opT, out int opA))
                    {
                        if (opP != projectId || opT != taskId || opA != attemptNo)
                            continue;
                    }
                    else if (curP != projectId || curT != taskId || curA != attemptNo)
                    {
                        continue;
                    }

                    string afterOp = line.Substring(opIdx + 4).Trim();
                    if (afterOp.Length > 0)
                        result.Add(afterOp);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] GetOperationsForTask: {ex.Message}");
            }

            return result;
        }

        public static bool HasDisallowedOperations(int projectId, int taskId, int attemptNo, HashSet<string> allowedOperationTypes)
        {
            if (allowedOperationTypes == null)
                return false;
            var ops = GetOperationsForTask(projectId, taskId, attemptNo);
            if (ops.Count == 0)
                return false;

            foreach (string opLine in ops)
            {
                string type = GetOperationType(opLine);
                if (string.IsNullOrEmpty(type))
                    continue;
                if (!allowedOperationTypes.Contains(type))
                    return true;
            }
            return false;
        }

        public static bool HasLoggedDestructiveError(int projectId, int taskId, int attemptNo)
        {
            try
            {
                string path = GetDestructiveErrorLogPath();
                if (!File.Exists(path))
                    return false;
                string prefix = $"{projectId},{taskId},{attemptNo}:";
                foreach (string line in ReadCachedLines(path))
                {
                    if (line != null && line.StartsWith(prefix, StringComparison.Ordinal))
                        return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] HasLoggedDestructiveError: {ex.Message}");
            }
            return false;
        }

        public static void AppendDestructiveErrors(int projectId, int taskId, int attemptNo, IList<string> errors)
        {
            if (errors == null || errors.Count == 0)
                return;
            try
            {
                string key = $"{projectId},{taskId},{attemptNo}:";
                string body = string.Join(" | ", errors.Where(e => !string.IsNullOrWhiteSpace(e)));
                File.AppendAllText(GetDestructiveErrorLogPath(), key + body + Environment.NewLine, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] AppendDestructiveErrors: {ex.Message}");
            }
            finally
            {
                InvalidateBatchScoringLogCache();
            }
        }

        /// <summary>VSTO の [Op] 行に付く [Task P-T-A] を解析（[Op] より前にある場合のみ）。</summary>
        private static bool TryParseExplicitOpTask(string line, int opIdx, out int projectId, out int taskId, out int attemptNo)
        {
            projectId = -1;
            taskId = -1;
            attemptNo = 0;
            int taskIdx = line.IndexOf("[Task ", StringComparison.OrdinalIgnoreCase);
            if (taskIdx < 0 || taskIdx >= opIdx)
                return false;
            int closeIdx = line.IndexOf(']', taskIdx);
            if (closeIdx < 0 || closeIdx > opIdx)
                return false;
            string inner = line.Substring(taskIdx + 6, closeIdx - taskIdx - 6).Trim();
            var tokens = inner.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length < 2
                || !int.TryParse(tokens[0].Trim(), out projectId)
                || !int.TryParse(tokens[1].Trim(), out taskId))
                return false;
            if (tokens.Length >= 3)
                int.TryParse(tokens[2].Trim(), out attemptNo);
            return projectId > 0 && taskId > 0;
        }

        private static void ParseTaskStart(string line, out int projectId, out int taskId, out int attemptNo)
        {
            projectId = -1;
            taskId = -1;
            attemptNo = 0;
            int startIdx = line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase);
            if (startIdx < 0)
                return;
            string part = line.Substring(startIdx + 11).Trim();
            var tokens = part.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length >= 2
                && int.TryParse(tokens[0].Trim(), out projectId)
                && int.TryParse(tokens[1].Trim(), out taskId))
            {
                if (tokens.Length >= 3)
                    int.TryParse(tokens[2].Trim(), out attemptNo);
            }
        }

        private static string GetOperationType(string opLine)
        {
            if (string.IsNullOrWhiteSpace(opLine))
                return "";
            int space = opLine.IndexOf(' ');
            return space > 0 ? opLine.Substring(0, space).Trim() : opLine.Trim();
        }

        private static string[] ReadCachedLines(string path)
        {
            if (!IsBatchCacheEnabled())
                return ReadLinesUncached(path);

            int generation = CurrentBatchCacheGeneration();
            bool destructive = IsDestructiveLogPath(path);
            if (TryReadFreshCache(path, destructive, generation, out string[] cached))
                return cached;

            if (!TryGetFileStamp(path, out bool existsBefore, out long lengthBefore, out long ticksBefore))
                return ReadLinesUncached(path);

            string[] lines = ReadLinesUncached(path);
            if (!TryGetFileStamp(path, out bool existsAfter, out long lengthAfter, out long ticksAfter))
                return lines;
            if (existsBefore != existsAfter || lengthBefore != lengthAfter || ticksBefore != ticksAfter)
                return lines;

            lock (CacheSync)
            {
                var stored = new CachedLines
                {
                    Generation = generation,
                    Exists = existsAfter,
                    Length = lengthAfter,
                    WriteTicks = ticksAfter,
                    Lines = lines
                };
                if (destructive)
                    DestructiveFileCache = stored;
                else
                    LogFileCache = stored;
            }
            return lines;
        }

        private static bool TryReadFreshCache(string path, bool destructive, int generation, out string[] lines)
        {
            lines = null;
            if (!TryGetFileStamp(path, out bool exists, out long length, out long ticks))
                return false;

            lock (CacheSync)
            {
                CachedLines cached = destructive ? DestructiveFileCache : LogFileCache;
                if (cached.Lines == null || cached.Generation != generation)
                    return false;
                if (cached.Exists != exists || cached.Length != length || cached.WriteTicks != ticks)
                    return false;
                lines = cached.Lines;
                return true;
            }
        }

        private static string[] ReadLinesUncached(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path) || !File.Exists(path))
                    return new string[0];
                return File.ReadAllLines(path, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] ReadLinesUncached: {ex.Message}");
                return new string[0];
            }
        }

        private static bool TryGetFileStamp(string path, out bool exists, out long length, out long writeTicks)
        {
            exists = false;
            length = 0;
            writeTicks = 0;
            try
            {
                if (string.IsNullOrEmpty(path) || !File.Exists(path))
                    return true;
                var info = new FileInfo(path);
                exists = true;
                length = info.Length;
                writeTicks = info.LastWriteTimeUtc.Ticks;
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] TryGetFileStamp: {ex.Message}");
                return false;
            }
        }

        private static bool IsBatchCacheEnabled()
        {
            return AppDomain.CurrentDomain.GetData(BatchCacheEnabledKey) as bool? == true;
        }

        private static int CurrentBatchCacheGeneration()
        {
            object current = AppDomain.CurrentDomain.GetData(BatchCacheGenerationKey);
            return current is int value ? value : 0;
        }

        private static bool IsDestructiveLogPath(string path)
        {
            return string.Equals(path, GetDestructiveErrorLogPath(), StringComparison.OrdinalIgnoreCase);
        }

        private struct CachedLines
        {
            public int Generation;
            public bool Exists;
            public long Length;
            public long WriteTicks;
            public string[] Lines;
        }
    }

    /// <summary>タスク単位の attempt 番号（通常 0、結果画面からの再採点は 1 以上）。</summary>
    public static class WordTaskAttemptRegistry
    {
        private static readonly object Sync = new object();
        private static readonly Dictionary<string, int> Attempts = new Dictionary<string, int>(StringComparer.Ordinal);

        private static string Key(int projectId, int taskId) => $"{projectId}-{taskId}";

        public static int GetAttempt(int projectId, int taskId)
        {
            lock (Sync)
            {
                if (Attempts.TryGetValue(Key(projectId, taskId), out int v))
                    return Math.Max(0, v);
                return 0;
            }
        }

        public static void SetAttempt(int projectId, int taskId, int attemptNo)
        {
            if (attemptNo < 0) attemptNo = 0;
            lock (Sync)
            {
                Attempts[Key(projectId, taskId)] = attemptNo;
            }
        }

        public static void ClearAll()
        {
            lock (Sync)
            {
                Attempts.Clear();
            }
        }

        public static void ClearProject(int projectId)
        {
            string prefix = projectId + "-";
            lock (Sync)
            {
                var keys = new List<string>();
                foreach (var k in Attempts.Keys)
                {
                    if (k.StartsWith(prefix, StringComparison.Ordinal))
                        keys.Add(k);
                }
                foreach (var k in keys)
                    Attempts.Remove(k);
            }
        }
    }
}
