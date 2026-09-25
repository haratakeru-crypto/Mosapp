using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MosPracticeClient;
using Newtonsoft.Json;

namespace MOSExcelMogiApp.Vocabulary
{
    public sealed class VocabularyKeywordItem
    {
        public string Id { get; set; }
        public string Category { get; set; }
        public string Keyword { get; set; }
        public string Answer { get; set; }
        public string Kind { get; set; }
        public string TargetTab { get; set; }
        public string TargetControl { get; set; }
        public List<string> DetectKeys { get; set; } = new List<string>();
        public string HighlightHint { get; set; }
        public string CoachMessage { get; set; }
        public string FormulaName { get; set; }
        public string Prefix3 { get; set; }

        public bool IsFunction =>
            string.Equals(Category, "Function", StringComparison.OrdinalIgnoreCase)
            || string.Equals(Kind, "Function", StringComparison.OrdinalIgnoreCase);
    }

    public sealed class VocabularyCatalogFile
    {
        public int Version { get; set; }
        public string TutorialTabKeywordId { get; set; }
        public string TutorialFunctionKeywordId { get; set; }
        public List<VocabularyKeywordItem> Items { get; set; } = new List<VocabularyKeywordItem>();
    }

    public static class VocabularyCatalog
    {
        static VocabularyCatalogFile _cache;

        public static VocabularyCatalogFile Load()
        {
            // 毎回最新 JSON を読む（開発中の差し替え反映）
            string path = ResolveJsonPath();
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("ExcelVocabularyKeywords.json が見つかりません。", path);

            string json = File.ReadAllText(path);
            _cache = JsonConvert.DeserializeObject<VocabularyCatalogFile>(json)
                     ?? new VocabularyCatalogFile();
            return _cache;
        }

        public static IReadOnlyList<VocabularyKeywordItem> Filter(VocabularyCategory category)
        {
            var catalog = Load();
            IEnumerable<VocabularyKeywordItem> items = catalog.Items ?? Enumerable.Empty<VocabularyKeywordItem>();
            switch (category)
            {
                case VocabularyCategory.TabButton:
                    items = items.Where(i => !i.IsFunction);
                    break;
                case VocabularyCategory.Function:
                    items = items.Where(i => i.IsFunction);
                    break;
                case VocabularyCategory.Both:
                    break;
                default:
                    items = Enumerable.Empty<VocabularyKeywordItem>();
                    break;
            }

            return items.ToList();
        }

        public static VocabularyKeywordItem FindById(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            return Load().Items?.FirstOrDefault(i =>
                string.Equals(i.Id, id, StringComparison.OrdinalIgnoreCase));
        }

        static string ResolveJsonPath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates =
            {
                Path.Combine(baseDir, "References", "JSON", "ExcelVocabularyKeywords.json"),
                Path.Combine(baseDir, "..", "..", "References", "JSON", "ExcelVocabularyKeywords.json"),
                Path.Combine(Directory.GetCurrentDirectory(), "References", "JSON", "ExcelVocabularyKeywords.json"),
            };
            return candidates.FirstOrDefault(File.Exists);
        }
    }
}
