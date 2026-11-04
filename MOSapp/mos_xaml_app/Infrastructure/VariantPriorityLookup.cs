using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace MOSExcelMogiApp.Infrastructure
{
    /// <summary>
    /// 類題結果画面用の優先度（CSV の ●/▲）。表示は優先/次点。採点には使わない。
    /// </summary>
    public static class VariantPriorityLookup
    {
        public const string FileNameSet1 = "Excel類題_セット1_優先度.csv";

        private static readonly object CacheGate = new object();
        private static string _cachedPath;
        private static DateTime _cachedWriteTimeUtc;
        private static Dictionary<string, string> _cachedMarks = new Dictionary<string, string>(StringComparer.Ordinal);

        public static string GetBadgeText(int variantSetNo, int projectId, int taskId)
        {
            if (variantSetNo != 1 || projectId <= 0 || taskId <= 0)
                return string.Empty;

            string mark = GetMark(projectId, taskId);
            if (mark == "●" || mark == "優先")
                return "優先";
            if (mark == "▲" || mark == "次点")
                return "次点";
            return string.Empty;
        }

        private static string GetMark(int projectId, int taskId)
        {
            Dictionary<string, string> marks = LoadMarks();
            string key = Key(projectId, taskId);
            return marks.TryGetValue(key, out string mark) ? mark : string.Empty;
        }

        private static Dictionary<string, string> LoadMarks()
        {
            string path = ResolvePath();
            DateTime writeTimeUtc = File.Exists(path)
                ? File.GetLastWriteTimeUtc(path)
                : DateTime.MinValue;

            lock (CacheGate)
            {
                if (string.Equals(_cachedPath, path, StringComparison.OrdinalIgnoreCase) &&
                    _cachedWriteTimeUtc == writeTimeUtc)
                {
                    return _cachedMarks;
                }

                _cachedPath = path;
                _cachedWriteTimeUtc = writeTimeUtc;
                _cachedMarks = ReadMarks(path);
                return _cachedMarks;
            }
        }

        private static string ResolvePath()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string referencesPath = Path.Combine(baseDirectory, "References", "CSV", FileNameSet1);
            if (File.Exists(referencesPath))
                return referencesPath;

            string flatPath = Path.Combine(baseDirectory, FileNameSet1);
            return File.Exists(flatPath) ? flatPath : referencesPath;
        }

        private static Dictionary<string, string> ReadMarks(string path)
        {
            var marks = new Dictionary<string, string>(StringComparer.Ordinal);
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                System.Diagnostics.Debug.WriteLine($"[VariantPriorityLookup] file not found: {path}");
                return marks;
            }

            try
            {
                foreach (string rawLine in File.ReadAllLines(path, new UTF8Encoding(true)))
                {
                    string line = (rawLine ?? string.Empty).Trim();
                    if (line.Length == 0)
                        continue;

                    string[] parts = line.Split(',');
                    if (parts.Length < 2)
                        continue;

                    string projectText = (parts[0] ?? string.Empty).Trim();
                    if (projectText == "プロジェクト")
                        continue;

                    if (!TryParseId(projectText, out int projectId) ||
                        !TryParseId((parts[1] ?? string.Empty).Trim(), out int taskId))
                    {
                        continue;
                    }

                    string mark = parts.Length > 2 ? (parts[2] ?? string.Empty).Trim() : string.Empty;
                    marks[Key(projectId, taskId)] = mark;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VariantPriorityLookup] read failed: {ex.Message}");
                marks.Clear();
            }

            return marks;
        }

        private static string Key(int projectId, int taskId)
        {
            return projectId.ToString(CultureInfo.InvariantCulture) + "-" +
                   taskId.ToString(CultureInfo.InvariantCulture);
        }

        private static bool TryParseId(string text, out int id)
        {
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out id) && id > 0;
        }
    }
}
