using System.Collections.Generic;

namespace MOSExcelMogiApp.Models
{
    /// <summary>
    /// 解答済み／あとで見直すを、教材と類題セット別にアプリ実行中保持する。
    /// AppBar を閉じても同じバケットを参照する。
    /// </summary>
    public static class TaskReviewStateStore
    {
        private static readonly Dictionary<string, Dictionary<int, bool[]>> CompletedByScope =
            new Dictionary<string, Dictionary<int, bool[]>>();
        private static readonly Dictionary<string, Dictionary<int, bool[]>> FlaggedByScope =
            new Dictionary<string, Dictionary<int, bool[]>>();

        public static Dictionary<int, bool[]> Completed(ExamResultScope scope)
        {
            return Bucket(CompletedByScope, scope);
        }

        public static Dictionary<int, bool[]> Flagged(ExamResultScope scope)
        {
            return Bucket(FlaggedByScope, scope);
        }

        public static void ClearProject(ExamResultScope scope, int projectId)
        {
            Completed(scope).Remove(projectId);
            Flagged(scope).Remove(projectId);
        }

        public static void Clear(ExamResultScope scope)
        {
            Completed(scope).Clear();
            Flagged(scope).Clear();
        }

        private static Dictionary<int, bool[]> Bucket(
            Dictionary<string, Dictionary<int, bool[]>> store,
            ExamResultScope scope)
        {
            string key = (scope ?? ExamResultScope.Textbook).Key;
            if (!store.TryGetValue(key, out Dictionary<int, bool[]> bucket) || bucket == null)
            {
                bucket = new Dictionary<int, bool[]>();
                store[key] = bucket;
            }

            return bucket;
        }
    }
}
