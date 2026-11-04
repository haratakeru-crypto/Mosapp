using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Libraries
{
    /// <summary>
    /// 採点結果をプロセス内で共有するためのストア。
    /// キーは教材 `{group}-{project}-{task}`、類題 `{group}-V{set}-{project}-{task}`。
    /// </summary>
    public static class ScoreResultStore
    {
        private static readonly object _lock = new object();
        private static readonly Dictionary<string, bool> _taskResults = new Dictionary<string, bool>(StringComparer.Ordinal);
        private static readonly Dictionary<string, string> _taskFailReasons = new Dictionary<string, string>(StringComparer.Ordinal);
        private static readonly HashSet<string> _incorrectTaskKeys = new HashSet<string>(StringComparer.Ordinal);
        private static readonly Dictionary<string, bool> _initialTaskResults = new Dictionary<string, bool>(StringComparer.Ordinal);
        private static readonly Dictionary<string, string> _initialTaskFailReasons = new Dictionary<string, string>(StringComparer.Ordinal);

        private static string MakeKey(int groupId, int projectId, int taskNumber, int variantSetNo)
        {
            if (variantSetNo > 0)
            {
                return groupId.ToString(CultureInfo.InvariantCulture)
                    + "-V" + variantSetNo.ToString(CultureInfo.InvariantCulture)
                    + "-" + projectId.ToString(CultureInfo.InvariantCulture)
                    + "-" + taskNumber.ToString(CultureInfo.InvariantCulture);
            }

            return groupId.ToString(CultureInfo.InvariantCulture)
                + "-" + projectId.ToString(CultureInfo.InvariantCulture)
                + "-" + taskNumber.ToString(CultureInfo.InvariantCulture);
        }

        private static bool MatchesScope(string key, int groupId, int variantSetNo)
        {
            if (string.IsNullOrEmpty(key))
                return false;

            string groupPrefix = groupId.ToString(CultureInfo.InvariantCulture) + "-";
            if (!key.StartsWith(groupPrefix, StringComparison.Ordinal))
                return false;

            string rest = key.Substring(groupPrefix.Length);
            if (variantSetNo > 0)
            {
                string variantPrefix = "V" + variantSetNo.ToString(CultureInfo.InvariantCulture) + "-";
                return rest.StartsWith(variantPrefix, StringComparison.Ordinal);
            }

            return rest.Length > 0 && rest[0] != 'V' && rest[0] != 'v';
        }

        private static bool TryParseProjectTask(string key, int groupId, int variantSetNo, out int projectId, out int taskNumber)
        {
            projectId = 0;
            taskNumber = 0;
            if (!MatchesScope(key, groupId, variantSetNo))
                return false;

            string groupPrefix = groupId.ToString(CultureInfo.InvariantCulture) + "-";
            string rest = key.Substring(groupPrefix.Length);
            if (variantSetNo > 0)
            {
                string variantPrefix = "V" + variantSetNo.ToString(CultureInfo.InvariantCulture) + "-";
                if (!rest.StartsWith(variantPrefix, StringComparison.Ordinal))
                    return false;
                rest = rest.Substring(variantPrefix.Length);
            }

            string[] parts = rest.Split('-');
            if (parts.Length != 2)
                return false;
            return int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out projectId)
                && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out taskNumber)
                && projectId > 0
                && taskNumber > 0;
        }

        /// <summary>指定グループ・スコープ（教材 / 類題セット）の採点結果をクリアする。</summary>
        public static void ClearGroup(int groupId, int variantSetNo = 0)
        {
            lock (_lock)
            {
                RemoveKeysMatching(_taskResults, groupId, variantSetNo);
                RemoveKeysMatching(_taskFailReasons, groupId, variantSetNo);
                _incorrectTaskKeys.RemoveWhere(k => MatchesScope(k, groupId, variantSetNo));
                RemoveKeysMatching(_initialTaskResults, groupId, variantSetNo);
                RemoveKeysMatching(_initialTaskFailReasons, groupId, variantSetNo);
            }
        }

        private static void RemoveKeysMatching<T>(Dictionary<string, T> map, int groupId, int variantSetNo)
        {
            foreach (var key in map.Keys.Where(k => MatchesScope(k, groupId, variantSetNo)).ToList())
                map.Remove(key);
        }

        /// <summary>現在の採点結果を初回スナップショットとして保存する。</summary>
        public static void SnapshotGroup(int groupId, int variantSetNo = 0)
        {
            lock (_lock)
            {
                RemoveKeysMatching(_initialTaskResults, groupId, variantSetNo);
                RemoveKeysMatching(_initialTaskFailReasons, groupId, variantSetNo);
                foreach (var kv in _taskResults.Where(k => MatchesScope(k.Key, groupId, variantSetNo)))
                    _initialTaskResults[kv.Key] = kv.Value;
                foreach (var kv in _taskFailReasons.Where(k => MatchesScope(k.Key, groupId, variantSetNo)))
                    _initialTaskFailReasons[kv.Key] = kv.Value;
            }
        }

        public static bool TryGetInitialResult(int groupId, int projectId, int taskNumber, out bool isPassed, int variantSetNo = 0)
        {
            var key = MakeKey(groupId, projectId, taskNumber, variantSetNo);
            lock (_lock)
            {
                if (_initialTaskResults.TryGetValue(key, out isPassed))
                    return true;
                isPassed = false;
                return false;
            }
        }

        public static string GetInitialFailReason(int groupId, int projectId, int taskNumber, int variantSetNo = 0)
        {
            var key = MakeKey(groupId, projectId, taskNumber, variantSetNo);
            lock (_lock)
            {
                if (_initialTaskFailReasons.TryGetValue(key, out string reason))
                    return reason ?? "";
                return "";
            }
        }

        /// <summary>初回採点で不正解だったタスクのキー一覧（"{projectId}-{taskId}" 形式）。</summary>
        public static IEnumerable<string> GetInitialWrongKeys(int groupId, int variantSetNo = 0)
        {
            lock (_lock)
            {
                foreach (var kv in _initialTaskResults)
                {
                    if (!MatchesScope(kv.Key, groupId, variantSetNo) || kv.Value)
                        continue;
                    if (_initialTaskFailReasons.TryGetValue(kv.Key, out string reason)
                        && string.Equals(reason, WordScoreExplanation.UnimplementedText, StringComparison.Ordinal))
                        continue;
                    if (!TryParseProjectTask(kv.Key, groupId, variantSetNo, out int projectId, out int taskNumber))
                        continue;
                    yield return projectId.ToString(CultureInfo.InvariantCulture)
                        + "-" + taskNumber.ToString(CultureInfo.InvariantCulture);
                }
            }
        }

        public static HashSet<int> GetScoredProjectIds(int groupId, int variantSetNo = 0)
        {
            return new HashSet<int>(GetProjectResults(groupId, variantSetNo).Keys);
        }

        public static Dictionary<int, List<bool>> GetProjectResults(int groupId, int variantSetNo = 0)
        {
            var byProject = new Dictionary<int, SortedDictionary<int, bool>>();
            lock (_lock)
            {
                foreach (var kv in _taskResults)
                {
                    if (!TryParseProjectTask(kv.Key, groupId, variantSetNo, out int projectId, out int taskNumber))
                        continue;
                    if (!byProject.TryGetValue(projectId, out var tasks))
                    {
                        tasks = new SortedDictionary<int, bool>();
                        byProject[projectId] = tasks;
                    }
                    tasks[taskNumber] = kv.Value;
                }
            }

            var results = new Dictionary<int, List<bool>>();
            foreach (var kv in byProject)
                results[kv.Key] = kv.Value.Values.ToList();
            return results;
        }

        public static void RecordResult(
            int groupId, int projectId, int taskNumber, bool isPassed, string failReason = null, int variantSetNo = 0)
        {
            var key = MakeKey(groupId, projectId, taskNumber, variantSetNo);
            lock (_lock)
            {
                _taskResults[key] = isPassed;
                _taskFailReasons[key] = isPassed ? "" : (failReason ?? "");
                bool unimplemented = !isPassed
                    && string.Equals(failReason, WordScoreExplanation.UnimplementedText, StringComparison.Ordinal);
                if (!isPassed && !unimplemented)
                    _incorrectTaskKeys.Add(key);
                else
                    _incorrectTaskKeys.Remove(key);
            }
        }

        public static string GetFailReason(int groupId, int projectId, int taskNumber, int variantSetNo = 0)
        {
            var key = MakeKey(groupId, projectId, taskNumber, variantSetNo);
            lock (_lock)
            {
                if (_taskFailReasons.TryGetValue(key, out string reason))
                    return reason ?? "";
                return "";
            }
        }

        public static bool IsScored(int groupId, int projectId, int taskNumber, int variantSetNo = 0)
        {
            var key = MakeKey(groupId, projectId, taskNumber, variantSetNo);
            lock (_lock)
            {
                return _taskResults.ContainsKey(key);
            }
        }

        public static bool TryGetResult(int groupId, int projectId, int taskNumber, out bool isPassed, int variantSetNo = 0)
        {
            var key = MakeKey(groupId, projectId, taskNumber, variantSetNo);
            lock (_lock)
            {
                if (_taskResults.TryGetValue(key, out isPassed))
                    return true;
                isPassed = false;
                return false;
            }
        }

        public static bool IsIncorrect(int groupId, int projectId, int taskNumber, int variantSetNo = 0)
        {
            var key = MakeKey(groupId, projectId, taskNumber, variantSetNo);
            lock (_lock)
            {
                if (_taskResults.TryGetValue(key, out bool isPassed))
                    return !isPassed;
                return _incorrectTaskKeys.Contains(key);
            }
        }
    }
}
