using System.Collections.Generic;

namespace MOSExcelMogiApp.Models
{
    /// <summary>
    /// 試験結果を保存するための静的ストレージクラス。
    /// 教材と類題セットは別バケットに保存する。
    /// </summary>
    public static class ExamResultStorage
    {
        public delegate void ProjectResultsChangedHandler(int projectId);

        /// <summary>
        /// 採点結果が更新されたときに発火（復習採点などでResultWindowを即時更新する用途）
        /// </summary>
        public static event ProjectResultsChangedHandler ResultsChanged;

        private static readonly Dictionary<string, Dictionary<int, List<bool>>> ResultsByScope =
            new Dictionary<string, Dictionary<int, List<bool>>>();
        private static readonly Dictionary<string, Dictionary<int, List<string>>> FailReasonsByScope =
            new Dictionary<string, Dictionary<int, List<string>>>();
        private static readonly Dictionary<string, Dictionary<int, List<bool>>> InitialResultsByScope =
            new Dictionary<string, Dictionary<int, List<bool>>>();

        public static void SaveProjectResult(int projectId, List<bool> results)
        {
            SaveProjectResult(ExamResultScope.Textbook, projectId, results, null);
        }

        public static void SaveProjectResult(int projectId, List<bool> results, IList<string> failReasons)
        {
            SaveProjectResult(ExamResultScope.Textbook, projectId, results, failReasons);
        }

        public static void SaveProjectResult(
            ExamResultScope scope,
            int projectId,
            List<bool> results,
            IList<string> failReasons)
        {
            Dictionary<int, List<bool>> projectResults = Bucket(ResultsByScope, scope);
            Dictionary<int, List<string>> projectFailReasons = Bucket(FailReasonsByScope, scope);

            projectResults[projectId] = results != null ? new List<bool>(results) : new List<bool>();
            if (failReasons == null)
                projectFailReasons.Remove(projectId);
            else
                projectFailReasons[projectId] = new List<string>(failReasons);

            try
            {
                ResultsChanged?.Invoke(projectId);
            }
            catch
            {
            }
        }

        public static string GetFailReason(int projectId, int taskId)
        {
            return GetFailReason(ExamResultScope.Textbook, projectId, taskId);
        }

        public static string GetFailReason(ExamResultScope scope, int projectId, int taskId)
        {
            if (taskId <= 0)
                return "";

            Dictionary<int, List<string>> reasonsByProject = Bucket(FailReasonsByScope, scope);
            if (!reasonsByProject.TryGetValue(projectId, out List<string> reasons) || reasons == null)
                return "";

            int index = taskId - 1;
            if (index < 0 || index >= reasons.Count || reasons[index] == null)
                return "";
            return reasons[index];
        }

        public static void SaveInitialResultsIfEmpty(Dictionary<int, List<bool>> allResults)
        {
            SaveInitialResultsIfEmpty(ExamResultScope.Textbook, allResults);
        }

        public static void SaveInitialResultsIfEmpty(ExamResultScope scope, Dictionary<int, List<bool>> allResults)
        {
            if (allResults == null || allResults.Count == 0)
                return;

            Dictionary<int, List<bool>> initial = Bucket(InitialResultsByScope, scope);
            if (initial.Count > 0)
                return;

            foreach (var kv in allResults)
                initial[kv.Key] = kv.Value != null ? new List<bool>(kv.Value) : new List<bool>();
        }

        public static Dictionary<int, List<bool>> GetAllResults()
        {
            return GetAllResults(ExamResultScope.Textbook);
        }

        public static Dictionary<int, List<bool>> GetAllResults(ExamResultScope scope)
        {
            return Clone(Bucket(ResultsByScope, scope));
        }

        public static Dictionary<int, List<bool>> GetInitialResults()
        {
            return GetInitialResults(ExamResultScope.Textbook);
        }

        public static Dictionary<int, List<bool>> GetInitialResults(ExamResultScope scope)
        {
            return Clone(Bucket(InitialResultsByScope, scope));
        }

        public static void Clear()
        {
            Clear(ExamResultScope.Textbook);
        }

        public static void Clear(ExamResultScope scope)
        {
            string key = (scope ?? ExamResultScope.Textbook).Key;
            if (ResultsByScope.TryGetValue(key, out Dictionary<int, List<bool>> results))
                results.Clear();
            if (FailReasonsByScope.TryGetValue(key, out Dictionary<int, List<string>> reasons))
                reasons.Clear();
            if (InitialResultsByScope.TryGetValue(key, out Dictionary<int, List<bool>> initial))
                initial.Clear();
        }

        private static Dictionary<int, T> Bucket<T>(
            Dictionary<string, Dictionary<int, T>> store,
            ExamResultScope scope)
        {
            string key = (scope ?? ExamResultScope.Textbook).Key;
            if (!store.TryGetValue(key, out Dictionary<int, T> bucket) || bucket == null)
            {
                bucket = new Dictionary<int, T>();
                store[key] = bucket;
            }

            return bucket;
        }

        private static Dictionary<int, List<bool>> Clone(Dictionary<int, List<bool>> source)
        {
            var clone = new Dictionary<int, List<bool>>();
            if (source == null)
                return clone;

            foreach (var kv in source)
                clone[kv.Key] = kv.Value != null ? new List<bool>(kv.Value) : new List<bool>();
            return clone;
        }
    }
}
