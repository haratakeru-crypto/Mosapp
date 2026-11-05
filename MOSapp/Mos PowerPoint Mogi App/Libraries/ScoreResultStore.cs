using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Libraries
{
    /// <summary>
    /// 採点結果をプロセス内で共有する。
    /// キーは教材 `{group}-{project}-{task}`、類題 `{group}-V{set}-{project}-{task}`。
    /// </summary>
    public static class ScoreResultStore
    {
        private static readonly object Sync = new object();
        private static readonly Dictionary<string, bool> TaskResults = new Dictionary<string, bool>(StringComparer.Ordinal);
        private static readonly Dictionary<string, string> TaskFailReasons = new Dictionary<string, string>(StringComparer.Ordinal);
        private static readonly Dictionary<string, bool> InitialTaskResults = new Dictionary<string, bool>(StringComparer.Ordinal);

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

        public static void ClearGroup(int groupId, int variantSetNo = 0)
        {
            lock (Sync)
            {
                RemoveKeysMatching(TaskResults, groupId, variantSetNo);
                RemoveKeysMatching(TaskFailReasons, groupId, variantSetNo);
                RemoveKeysMatching(InitialTaskResults, groupId, variantSetNo);
            }
        }

        private static void RemoveKeysMatching<T>(Dictionary<string, T> map, int groupId, int variantSetNo)
        {
            foreach (var key in map.Keys.Where(k => MatchesScope(k, groupId, variantSetNo)).ToList())
                map.Remove(key);
        }

        public static void SnapshotGroupIfEmpty(int groupId, int variantSetNo = 0)
        {
            lock (Sync)
            {
                if (InitialTaskResults.Keys.Any(k => MatchesScope(k, groupId, variantSetNo)))
                    return;

                foreach (var kv in TaskResults.Where(k => MatchesScope(k.Key, groupId, variantSetNo)))
                    InitialTaskResults[kv.Key] = kv.Value;
            }
        }

        public static void RecordResult(
            int groupId, int projectId, int taskNumber, bool isPassed, string failReason = null, int variantSetNo = 0)
        {
            var key = MakeKey(groupId, projectId, taskNumber, variantSetNo);
            lock (Sync)
            {
                TaskResults[key] = isPassed;
                TaskFailReasons[key] = isPassed ? "" : (failReason ?? "");
            }
        }

        public static void RecordProjectResults(
            int groupId, int projectId, IList<bool> results, int variantSetNo = 0)
        {
            if (results == null)
                return;

            for (int i = 0; i < results.Count; i++)
                RecordResult(groupId, projectId, i + 1, results[i], null, variantSetNo);
        }

        public static Dictionary<int, List<bool>> GetProjectResults(int groupId, int variantSetNo = 0)
        {
            var byProject = new Dictionary<int, SortedDictionary<int, bool>>();
            lock (Sync)
            {
                foreach (var kv in TaskResults)
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

        public static HashSet<int> GetScoredProjectIds(int groupId, int variantSetNo = 0)
        {
            return new HashSet<int>(GetProjectResults(groupId, variantSetNo).Keys);
        }
    }
}
