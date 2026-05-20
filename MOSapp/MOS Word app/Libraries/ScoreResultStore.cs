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
