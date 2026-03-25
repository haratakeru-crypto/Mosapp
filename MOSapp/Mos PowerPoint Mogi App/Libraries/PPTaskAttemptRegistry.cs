using System;
using System.Collections.Generic;

namespace Libraries
{
    /// <summary>
    /// タスク単位の再挑戦 attempt 番号をプロセス内で共有する。
    /// UiTestAppBarWindow の再挑戦と、全件採点・単体採点で同じ attempt を参照するために使用する。
    /// </summary>
    public static class PPTaskAttemptRegistry
    {
        private static readonly object Sync = new object();
        private static readonly Dictionary<string, int> Attempts = new Dictionary<string, int>(StringComparer.Ordinal);

        private static string Key(int projectId, int taskId) => $"{projectId}-{taskId}";

        /// <summary>未登録時は 1 を返す。</summary>
        public static int GetAttempt(int projectId, int taskId)
        {
            lock (Sync)
            {
                if (Attempts.TryGetValue(Key(projectId, taskId), out int v) && v >= 1)
                    return v;
                return 1;
            }
        }

        public static void SetAttempt(int projectId, int taskId, int attemptNo)
        {
            if (attemptNo < 1) attemptNo = 1;
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

        /// <summary>単体プロジェクトリセット時、そのプロジェクトに属する attempt のみ削除する。</summary>
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
