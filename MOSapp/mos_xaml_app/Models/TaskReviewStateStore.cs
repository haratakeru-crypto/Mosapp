using System;
using System.Collections.Generic;
using System.IO;
using MosPracticeClient;
using Newtonsoft.Json;

namespace MOSExcelMogiApp.Models
{
    /// <summary>
    /// 解答済み／あとで見直すを、教材と類題セット別に保持する。
    /// アプリ終了後もローカル JSON に残し、リセットと受験者削除で消す。
    /// </summary>
    public static class TaskReviewStateStore
    {
        private static readonly Dictionary<string, Dictionary<int, bool[]>> CompletedByScope =
            new Dictionary<string, Dictionary<int, bool[]>>();
        private static readonly Dictionary<string, Dictionary<int, bool[]>> FlaggedByScope =
            new Dictionary<string, Dictionary<int, bool[]>>();
        private static bool _loaded;

        public static Dictionary<int, bool[]> Completed(ExamResultScope scope)
        {
            EnsureLoaded();
            return Bucket(CompletedByScope, scope);
        }

        public static Dictionary<int, bool[]> Flagged(ExamResultScope scope)
        {
            EnsureLoaded();
            return Bucket(FlaggedByScope, scope);
        }

        public static void ClearProject(ExamResultScope scope, int projectId)
        {
            Completed(scope).Remove(projectId);
            Flagged(scope).Remove(projectId);
            Persist();
        }

        public static void Clear(ExamResultScope scope)
        {
            Completed(scope).Clear();
            Flagged(scope).Clear();
            Persist();
        }

        public static void Persist()
        {
            EnsureLoaded();
            try
            {
                Directory.CreateDirectory(ExamineeStore.DataDirectory);
                var file = new TaskReviewStateFile
                {
                    UserKey = CurrentUserKey(),
                    Scopes = new Dictionary<string, TaskReviewScopeFile>()
                };

                foreach (var kv in CompletedByScope)
                    ScopeFile(file, kv.Key).Completed = ToLists(kv.Value);
                foreach (var kv in FlaggedByScope)
                    ScopeFile(file, kv.Key).Flagged = ToLists(kv.Value);

                File.WriteAllText(
                    ExamineeStore.ExcelTaskReviewStatePath,
                    JsonConvert.SerializeObject(file, Formatting.Indented));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[TaskReviewStateStore] Persist: " + ex.Message);
            }
        }

        private static void EnsureLoaded()
        {
            if (_loaded)
                return;
            _loaded = true;

            try
            {
                string path = ExamineeStore.ExcelTaskReviewStatePath;
                if (!File.Exists(path))
                    return;

                var file = JsonConvert.DeserializeObject<TaskReviewStateFile>(File.ReadAllText(path));
                if (file?.Scopes == null)
                    return;
                if (!string.Equals(file.UserKey ?? "", CurrentUserKey(), StringComparison.Ordinal))
                    return;

                foreach (var kv in file.Scopes)
                {
                    if (string.IsNullOrEmpty(kv.Key) || kv.Value == null)
                        continue;
                    CompletedByScope[kv.Key] = ToArrays(kv.Value.Completed);
                    FlaggedByScope[kv.Key] = ToArrays(kv.Value.Flagged);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[TaskReviewStateStore] Load: " + ex.Message);
            }
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

        private static string CurrentUserKey()
        {
            return ExamineeStore.VocabUserKey(ExamineeStore.Load()) ?? "";
        }

        private static TaskReviewScopeFile ScopeFile(TaskReviewStateFile file, string key)
        {
            if (!file.Scopes.TryGetValue(key, out TaskReviewScopeFile scope) || scope == null)
            {
                scope = new TaskReviewScopeFile();
                file.Scopes[key] = scope;
            }

            return scope;
        }

        private static Dictionary<int, List<bool>> ToLists(Dictionary<int, bool[]> source)
        {
            var clone = new Dictionary<int, List<bool>>();
            if (source == null)
                return clone;
            foreach (var kv in source)
                clone[kv.Key] = kv.Value != null ? new List<bool>(kv.Value) : new List<bool>();
            return clone;
        }

        private static Dictionary<int, bool[]> ToArrays(Dictionary<int, List<bool>> source)
        {
            var clone = new Dictionary<int, bool[]>();
            if (source == null)
                return clone;
            foreach (var kv in source)
                clone[kv.Key] = kv.Value != null ? kv.Value.ToArray() : Array.Empty<bool>();
            return clone;
        }

        private sealed class TaskReviewStateFile
        {
            public string UserKey { get; set; }
            public Dictionary<string, TaskReviewScopeFile> Scopes { get; set; }
        }

        private sealed class TaskReviewScopeFile
        {
            public Dictionary<int, List<bool>> Completed { get; set; }
            public Dictionary<int, List<bool>> Flagged { get; set; }
        }
    }
}
