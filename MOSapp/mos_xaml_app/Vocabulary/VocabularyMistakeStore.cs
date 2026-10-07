using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MosPracticeClient;
using Newtonsoft.Json;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>
    /// 単語帳で間違えた問題を、大学名と名前の組ごとに保存する。
    /// 未登録のときは何も保存しない。
    /// </summary>
    public static class VocabularyMistakeStore
    {
        public sealed class Entry
        {
            public string Keyword { get; set; }
            public bool IsFunction { get; set; }
        }

        static string CurrentUserKey()
        {
            return ExamineeStore.VocabUserKey(ExamineeStore.Load());
        }

        static Dictionary<string, List<Entry>> LoadAll()
        {
            try
            {
                if (!File.Exists(ExamineeStore.VocabMistakesPath))
                    return new Dictionary<string, List<Entry>>();
                return JsonConvert.DeserializeObject<Dictionary<string, List<Entry>>>(
                           File.ReadAllText(ExamineeStore.VocabMistakesPath))
                       ?? new Dictionary<string, List<Entry>>();
            }
            catch
            {
                return new Dictionary<string, List<Entry>>();
            }
        }

        static void SaveAll(Dictionary<string, List<Entry>> all)
        {
            try
            {
                Directory.CreateDirectory(ExamineeStore.DataDirectory);
                File.WriteAllText(ExamineeStore.VocabMistakesPath, JsonConvert.SerializeObject(all, Formatting.Indented));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[VocabularyMistakeStore] Save: " + ex.Message);
            }
        }

        public static void Record(VocabularyKeywordItem item)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.Keyword)) return;
            string user = CurrentUserKey();
            if (user == null) return;

            var all = LoadAll();
            if (!all.TryGetValue(user, out var list) || list == null)
            {
                list = new List<Entry>();
                all[user] = list;
            }
            if (list.Any(e => Same(e, item))) return;
            list.Add(new Entry { Keyword = item.Keyword.Trim(), IsFunction = item.IsFunction });
            SaveAll(all);
        }

        public static void Remove(VocabularyKeywordItem item)
        {
            if (item == null) return;
            string user = CurrentUserKey();
            if (user == null) return;

            var all = LoadAll();
            if (!all.TryGetValue(user, out var list) || list == null) return;
            if (list.RemoveAll(e => Same(e, item)) == 0) return;
            if (list.Count == 0) all.Remove(user);
            SaveAll(all);
        }

        /// <summary>今の利用者の、カテゴリに合う間違い問題。CSV の順で返す。</summary>
        public static List<VocabularyKeywordItem> LoadItems(VocabularyCategory category)
        {
            var result = new List<VocabularyKeywordItem>();
            string user = CurrentUserKey();
            if (user == null) return result;

            var all = LoadAll();
            if (!all.TryGetValue(user, out var list) || list == null || list.Count == 0) return result;

            foreach (var item in VocabularyCatalog.Filter(category))
            {
                if (list.Any(e => Same(e, item)))
                    result.Add(item);
            }
            return result;
        }

        static bool Same(Entry e, VocabularyKeywordItem item)
        {
            return e != null
                   && e.IsFunction == item.IsFunction
                   && string.Equals((e.Keyword ?? "").Trim(), (item.Keyword ?? "").Trim(), StringComparison.Ordinal);
        }
    }
}
