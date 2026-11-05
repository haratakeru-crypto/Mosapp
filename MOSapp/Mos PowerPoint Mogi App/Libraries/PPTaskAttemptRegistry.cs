using System;
using System.Collections.Generic;

namespace Libraries
{
    /// <summary>
    /// タスク単位の再挑戦 attempt 番号をプロセス内で共有する。
    /// UiTestAppBarWindow の再挑戦と、全件採点・単体採点で同じ attempt を参照するために使用する。
    /// 教材は <c>{project}-{task}</c>、類題は <c>V{set}-{project}-{task}</c>（Word/Excel と同型）。
    /// </summary>
    public static class PPTaskAttemptRegistry
    {
        private static readonly object Sync = new object();
        private static readonly Dictionary<string, int> Attempts = new Dictionary<string, int>(StringComparer.Ordinal);

        private static string Key(int projectId, int taskId, int variantSetNo)
        {
            if (variantSetNo > 0)
                return "V" + variantSetNo + "-" + projectId + "-" + taskId;
            return projectId + "-" + taskId;
        }

        /// <summary>未登録時は 1 を返す。</summary>
        public static int GetAttempt(int projectId, int taskId)
        {
            return GetAttempt(projectId, taskId, 0);
        }

        /// <summary>未登録時は 1 を返す。類題は variantSetNo &gt; 0。</summary>
        public static int GetAttempt(int projectId, int taskId, int variantSetNo)
        {
            lock (Sync)
            {
                if (Attempts.TryGetValue(Key(projectId, taskId, variantSetNo), out int v) && v >= 1)
                    return v;
                return 1;
            }
        }

        public static void SetAttempt(int projectId, int taskId, int attemptNo)
        {
            SetAttempt(projectId, taskId, attemptNo, 0);
        }

        public static void SetAttempt(int projectId, int taskId, int attemptNo, int variantSetNo)
        {
            if (attemptNo < 1) attemptNo = 1;
            lock (Sync)
            {
                Attempts[Key(projectId, taskId, variantSetNo)] = attemptNo;
            }
        }

        public static void ClearAll()
        {
            lock (Sync)
            {
                Attempts.Clear();
            }
        }

        /// <summary>単体プロジェクトリセット時、そのプロジェクトに属する attempt のみ削除する（教材 scope）。</summary>
        public static void ClearProject(int projectId)
        {
            ClearProject(projectId, 0);
        }

        /// <summary>教材 / 類題 scope ごとに attempt を削除する。</summary>
        public static void ClearProject(int projectId, int variantSetNo)
        {
            lock (Sync)
            {
                var keys = new List<string>();
                if (variantSetNo > 0)
                {
                    string prefix = "V" + variantSetNo + "-" + projectId + "-";
                    foreach (var k in Attempts.Keys)
                    {
                        if (k.StartsWith(prefix, StringComparison.Ordinal))
                            keys.Add(k);
                    }
                }
                else
                {
                    string prefix = projectId + "-";
                    foreach (var k in Attempts.Keys)
                    {
                        // 類題キー Vn-P-T は除外
                        if (k.Length > 0 && (k[0] == 'V' || k[0] == 'v'))
                            continue;
                        if (k.StartsWith(prefix, StringComparison.Ordinal))
                            keys.Add(k);
                    }
                }

                foreach (var k in keys)
                    Attempts.Remove(k);
            }
        }
    }
}
