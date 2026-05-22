using System;
using System.Collections.Generic;
using System.Linq;

namespace Libraries
{
    /// <summary>
    /// 採点結果をプロセス内で共有するためのストア。
    /// グループID＋プロジェクトID＋タスク番号をキーに、全問の正誤を保持する。
    /// </summary>
    public static class ScoreResultStore
    {
        private static readonly object _lock = new object();
        // key 形式: "{groupId}-{projectId}-{taskNumber}"
        private static readonly Dictionary<string, bool> _taskResults = new Dictionary<string, bool>(StringComparer.Ordinal);
        private static readonly HashSet<string> _incorrectTaskKeys = new HashSet<string>(StringComparer.Ordinal);
        /// <summary>試験終了直後の初回採点スナップショット（キーは MakeKey と同形式）</summary>
        private static readonly Dictionary<string, bool> _initialTaskResults = new Dictionary<string, bool>(StringComparer.Ordinal);

        private static string MakeKey(int groupId, int projectId, int taskNumber)
        {
            return $"{groupId}-{projectId}-{taskNumber}";
        }

        /// <summary>
        /// 指定グループの採点結果をクリアする（一括採点開始前に呼ぶ）。
        /// </summary>
        public static void ClearGroup(int groupId)
        {
            string prefix = $"{groupId}-";
            lock (_lock)
            {
                foreach (var key in _taskResults.Keys.Where(k => k.StartsWith(prefix, StringComparison.Ordinal)).ToList())
                    _taskResults.Remove(key);
                _incorrectTaskKeys.RemoveWhere(k => k.StartsWith(prefix, StringComparison.Ordinal));
                foreach (var key in _initialTaskResults.Keys.Where(k => k.StartsWith(prefix, StringComparison.Ordinal)).ToList())
                    _initialTaskResults.Remove(key);
            }
        }

        /// <summary>
        /// 現在の採点結果を初回スナップショットとして保存する（一括採点直後に1回呼ぶ）。
        /// </summary>
        public static void SnapshotGroup(int groupId)
        {
            string prefix = $"{groupId}-";
            lock (_lock)
            {
                foreach (var key in _initialTaskResults.Keys.Where(k => k.StartsWith(prefix, StringComparison.Ordinal)).ToList())
                    _initialTaskResults.Remove(key);
                foreach (var kv in _taskResults.Where(k => k.Key.StartsWith(prefix, StringComparison.Ordinal)))
                    _initialTaskResults[kv.Key] = kv.Value;
            }
        }

        /// <summary>
        /// 初回スナップショットから採点結果を取得する。
        /// </summary>
        public static bool TryGetInitialResult(int groupId, int projectId, int taskNumber, out bool isPassed)
        {
            var key = MakeKey(groupId, projectId, taskNumber);
            lock (_lock)
            {
                if (_initialTaskResults.TryGetValue(key, out isPassed))
                    return true;
                isPassed = false;
                return false;
            }
        }

        /// <summary>
        /// 初回採点で不正解だったタスクのキー一覧（"{projectId}-{taskId}" 形式）。
        /// </summary>
        public static IEnumerable<string> GetInitialWrongKeys(int groupId)
        {
            string prefix = $"{groupId}-";
            lock (_lock)
            {
                foreach (var kv in _initialTaskResults)
                {
                    if (!kv.Key.StartsWith(prefix, StringComparison.Ordinal) || kv.Value)
                        continue;
                    var parts = kv.Key.Substring(prefix.Length).Split('-');
                    if (parts.Length == 2)
                        yield return $"{parts[0]}-{parts[1]}";
                }
            }
        }

        /// <summary>
        /// 1 問分の採点結果を記録する。
        /// </summary>
        public static void RecordResult(int groupId, int projectId, int taskNumber, bool isPassed)
        {
            var key = MakeKey(groupId, projectId, taskNumber);
            lock (_lock)
            {
                _taskResults[key] = isPassed;
                if (!isPassed)
                    _incorrectTaskKeys.Add(key);
                else
                    _incorrectTaskKeys.Remove(key);
            }
        }

        /// <summary>
        /// 採点済みかどうかを返す。
        /// </summary>
        public static bool IsScored(int groupId, int projectId, int taskNumber)
        {
            var key = MakeKey(groupId, projectId, taskNumber);
            lock (_lock)
            {
                return _taskResults.ContainsKey(key);
            }
        }

        /// <summary>
        /// 採点結果を取得する。未採点の場合は false を返し、取得失敗時は false。
        /// </summary>
        public static bool TryGetResult(int groupId, int projectId, int taskNumber, out bool isPassed)
        {
            var key = MakeKey(groupId, projectId, taskNumber);
            lock (_lock)
            {
                if (_taskResults.TryGetValue(key, out isPassed))
                    return true;
                isPassed = false;
                return false;
            }
        }

        /// <summary>
        /// 指定の問題が「採点で不正解だったか」を取得する。
        /// </summary>
        public static bool IsIncorrect(int groupId, int projectId, int taskNumber)
        {
            var key = MakeKey(groupId, projectId, taskNumber);
            lock (_lock)
            {
                if (_taskResults.TryGetValue(key, out bool isPassed))
                    return !isPassed;
                return _incorrectTaskKeys.Contains(key);
            }
        }
    }
}
