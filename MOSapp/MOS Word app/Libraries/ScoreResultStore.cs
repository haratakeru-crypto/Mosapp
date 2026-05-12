using System;
using System.Collections.Generic;

namespace Libraries
{
    /// <summary>
    /// 採点結果（間違えた問題）をプロセス内で共有するためのストア。
    /// グループID＋プロジェクトID＋タスク番号をキーに、不正解だった問題を保持する。
    /// </summary>
    public static class ScoreResultStore
    {
        private static readonly object _lock = new object();
        // key 形式: "{groupId}-{projectId}-{taskNumber}"
        private static readonly HashSet<string> _incorrectTaskKeys = new HashSet<string>(StringComparer.Ordinal);

        private static string MakeKey(int groupId, int projectId, int taskNumber)
        {
            return $"{groupId}-{projectId}-{taskNumber}";
        }

        /// <summary>
        /// 1 問分の採点結果を記録する。
        /// </summary>
        public static void RecordResult(int groupId, int projectId, int taskNumber, bool isPassed)
        {
            var key = MakeKey(groupId, projectId, taskNumber);
            lock (_lock)
            {
                if (!isPassed)
                {
                    _incorrectTaskKeys.Add(key);
                }
                else
                {
                    _incorrectTaskKeys.Remove(key);
                }
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
                return _incorrectTaskKeys.Contains(key);
            }
        }
    }
}

