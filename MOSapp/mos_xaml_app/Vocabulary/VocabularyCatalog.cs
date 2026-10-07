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
        /// <summary>問題カードに出す文。CSV の表示テキスト。</summary>
        public string DisplayText { get; set; }

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
            int minLine;
            int maxLine;
            switch (category)
            {
                case VocabularyCategory.TabButton:
                    minLine = 2;
                    maxLine = 18;
                    break;
                case VocabularyCategory.Function:
                    minLine = 19;
                    maxLine = 37;
                    break;
                case VocabularyCategory.Both:
                    minLine = 2;
                    maxLine = 37;
                    break;
                default:
                    return new List<VocabularyKeywordItem>();
            }

            var templates = Load().Items ?? new List<VocabularyKeywordItem>();
            var rows = LoadCsvRows().Where(r => r.StartLine >= minLine && r.StartLine <= maxLine);
            var list = new List<VocabularyKeywordItem>();
            foreach (var row in rows)
            {
                var template = FindTemplate(templates, row);
                list.Add(CloneForRow(template, row, category));
            }
            return list;
        }

        sealed class CsvKeywordRow
        {
            public int StartLine;
            public string Keyword;
            public string Answer;
            public string DisplayText;
        }

        static VocabularyKeywordItem FindTemplate(List<VocabularyKeywordItem> templates, CsvKeywordRow row)
        {
            var exact = templates.FirstOrDefault(i =>
                string.Equals(i.Keyword, row.Keyword, StringComparison.Ordinal));
            if (exact != null) return exact;

            var byAnswer = templates.FirstOrDefault(i =>
                string.Equals(i.FormulaName, row.Answer, StringComparison.OrdinalIgnoreCase)
                || string.Equals(i.Answer, row.Answer, StringComparison.OrdinalIgnoreCase));
            if (byAnswer != null) return byAnswer;

            return templates.FirstOrDefault(i =>
                !string.IsNullOrEmpty(i.Keyword)
                && i.Keyword.IndexOf(row.Keyword ?? "", StringComparison.Ordinal) >= 0);
        }

        static VocabularyKeywordItem CloneForRow(VocabularyKeywordItem template, CsvKeywordRow row, VocabularyCategory category)
        {
            bool functionRow = row.StartLine >= 19;
            var item = new VocabularyKeywordItem
            {
                Id = template?.Id,
                Category = template?.Category ?? (functionRow ? "Function" : "TabButton"),
                Keyword = row.Keyword,
                Answer = string.IsNullOrWhiteSpace(row.Answer) ? template?.Answer : row.Answer,
                Kind = template?.Kind ?? (functionRow ? "Function" : "Button"),
                TargetTab = template?.TargetTab,
                TargetControl = template?.TargetControl,
                DetectKeys = template?.DetectKeys != null
                    ? new List<string>(template.DetectKeys)
                    : new List<string>(),
                HighlightHint = template?.HighlightHint ?? (functionRow ? "FormulaBar" : null),
                CoachMessage = template?.CoachMessage,
                FormulaName = template?.FormulaName,
                Prefix3 = template?.Prefix3,
                DisplayText = NormalizeDisplay(row.DisplayText)
            };

            if (functionRow && item.DetectKeys.Count == 0 && !string.IsNullOrWhiteSpace(item.Answer))
                item.DetectKeys.Add("Formula:" + item.Answer.Trim());
            if (functionRow && string.IsNullOrWhiteSpace(item.FormulaName))
                item.FormulaName = item.Answer;
            if (category == VocabularyCategory.Function)
                item.Category = "Function";
            else if (category == VocabularyCategory.TabButton)
                item.Category = "TabButton";
            return item;
        }

        static string NormalizeDisplay(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            return text.Replace("\r\n", "\n").Replace('\r', '\n').Trim();
        }

        static List<CsvKeywordRow> LoadCsvRows()
        {
            string path = ResolveCsvPath();
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                return new List<CsvKeywordRow>();
            return ParseCsv(File.ReadAllText(path));
        }

        static string ResolveCsvPath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates =
            {
                Path.Combine(baseDir, "References", "CSV", "ExcelVocabularyKeywords.csv"),
                Path.Combine(baseDir, "..", "..", "References", "CSV", "ExcelVocabularyKeywords.csv"),
                Path.Combine(Directory.GetCurrentDirectory(), "References", "CSV", "ExcelVocabularyKeywords.csv"),
            };
            return candidates.FirstOrDefault(File.Exists);
        }

        static List<CsvKeywordRow> ParseCsv(string text)
        {
            var rows = new List<CsvKeywordRow>();
            var field = new System.Text.StringBuilder();
            var fields = new List<string>();
            int line = 1;
            int recordStart = 1;
            bool inQuotes = false;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"')
                        {
                            field.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        if (c == '\n') line++;
                        else if (c == '\r')
                        {
                            line++;
                            if (i + 1 < text.Length && text[i + 1] == '\n') i++;
                        }
                        field.Append(c == '\r' ? '\n' : c);
                    }
                    continue;
                }

                if (c == '"')
                {
                    inQuotes = true;
                }
                else if (c == ',')
                {
                    fields.Add(field.ToString());
                    field.Clear();
                }
                else if (c == '\n' || c == '\r')
                {
                    if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    fields.Add(field.ToString());
                    field.Clear();
                    AddRow(rows, recordStart, fields);
                    fields.Clear();
                    line++;
                    recordStart = line;
                }
                else
                {
                    field.Append(c);
                }
            }

            if (field.Length > 0 || fields.Count > 0)
            {
                fields.Add(field.ToString());
                AddRow(rows, recordStart, fields);
            }

            if (rows.Count > 0 && string.Equals(rows[0].Keyword, "キーワード", StringComparison.Ordinal))
                rows.RemoveAt(0);
            return rows;
        }

        static void AddRow(List<CsvKeywordRow> rows, int startLine, List<string> fields)
        {
            if (fields.All(f => string.IsNullOrWhiteSpace(f))) return;
            rows.Add(new CsvKeywordRow
            {
                StartLine = startLine,
                Keyword = fields.Count > 0 ? fields[0].Trim() : "",
                Answer = fields.Count > 1 ? fields[1].Trim() : "",
                DisplayText = fields.Count > 2 ? fields[2].Trim() : ""
            });
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
