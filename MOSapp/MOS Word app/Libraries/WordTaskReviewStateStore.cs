using System;
using System.Collections.Generic;
using System.IO;
using MosPracticeClient;
using Newtonsoft.Json;

namespace MOS_Word_app
{
    /// <summary>
    /// 類題の解答済み／あとで見直すをローカル JSON に残す。
    /// アプリ終了後も残し、リセット・大学情報削除・Windows シャットダウンで消す。
    /// </summary>
    public static class WordTaskReviewStateStore
    {
        private static readonly Dictionary<int, bool[]> CompletedStates = new Dictionary<int, bool[]>();
        private static readonly Dictionary<int, bool[]> FlaggedStates = new Dictionary<int, bool[]>();
        private static bool _loaded;

        public static Dictionary<int, bool[]> Completed()
        {
            EnsureLoaded();
            return CompletedStates;
        }

        public static Dictionary<int, bool[]> Flagged()
        {
            EnsureLoaded();
            return FlaggedStates;
        }

        public static void ClearProject(int projectId)
        {
            Completed().Remove(projectId);
            Flagged().Remove(projectId);
            Persist();
        }

        public static void Clear()
        {
            Completed().Clear();
            Flagged().Clear();
            Persist();
        }

        public static void Discard()
        {
            CompletedStates.Clear();
            FlaggedStates.Clear();
            _loaded = true;
        }

        public static void Persist()
        {
            EnsureLoaded();
            try
            {
                Directory.CreateDirectory(ExamineeStore.DataDirectory);
                var file = new WordTaskReviewStateFile
                {
                    UserKey = CurrentUserKey(),
                    Completed = ToLists(CompletedStates),
                    Flagged = ToLists(FlaggedStates)
                };
                File.WriteAllText(
                    ExamineeStore.WordTaskReviewStatePath,
                    JsonConvert.SerializeObject(file, Formatting.Indented));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[WordTaskReviewStateStore] Persist: " + ex.Message);
            }
        }

        private static void EnsureLoaded()
        {
            if (_loaded)
                return;
            _loaded = true;

            try
            {
                string path = ExamineeStore.WordTaskReviewStatePath;
                if (!File.Exists(path))
                    return;

                var file = JsonConvert.DeserializeObject<WordTaskReviewStateFile>(File.ReadAllText(path));
                if (file == null)
                    return;
                if (!string.Equals(file.UserKey ?? "", CurrentUserKey(), StringComparison.Ordinal))
                    return;

                CopyInto(CompletedStates, ToArrays(file.Completed));
                CopyInto(FlaggedStates, ToArrays(file.Flagged));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[WordTaskReviewStateStore] Load: " + ex.Message);
            }
        }

        private static void CopyInto(Dictionary<int, bool[]> target, Dictionary<int, bool[]> source)
        {
            target.Clear();
            if (source == null)
                return;
            foreach (var kv in source)
                target[kv.Key] = kv.Value;
        }

        private static string CurrentUserKey()
        {
            return ExamineeStore.VocabUserKey(ExamineeStore.Load()) ?? "";
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

        private sealed class WordTaskReviewStateFile
        {
            public string UserKey { get; set; }
            public Dictionary<int, List<bool>> Completed { get; set; }
            public Dictionary<int, List<bool>> Flagged { get; set; }
        }
    }
}
