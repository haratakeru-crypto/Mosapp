using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace MosPracticeClient
{
    public class ScoringLogEntry
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("rangeLabel")]
        public string RangeLabel { get; set; }

        [JsonProperty("scoredAt")]
        public DateTime ScoredAt { get; set; }

        [JsonProperty("groupId")]
        public int GroupId { get; set; }

        [JsonProperty("isVariantMode")]
        public bool IsVariantMode { get; set; }

        [JsonProperty("variantSetNo")]
        public int VariantSetNo { get; set; }

        [JsonProperty("projectResults")]
        public Dictionary<int, List<bool>> ProjectResults { get; set; }

        [JsonProperty("completedStates")]
        public Dictionary<int, List<bool>> CompletedStates { get; set; }

        [JsonProperty("flaggedStates")]
        public Dictionary<int, List<bool>> FlaggedStates { get; set; }

        [JsonIgnore]
        public string DisplayTitle
        {
            get
            {
                string label = string.IsNullOrWhiteSpace(RangeLabel) ? "採点" : RangeLabel.Trim();
                if (IsVariantMode)
                    return $"類題 {label} {ScoredAt:M/d H:mm}";
                return $"{label} {ScoredAt:M/d H:mm}";
            }
        }

        public static ScoringLogEntry Create(
            string rangeLabel,
            int groupId,
            Dictionary<int, List<bool>> projectResults,
            bool isVariantMode = false,
            int variantSetNo = 0,
            Dictionary<int, bool[]> completedStates = null,
            Dictionary<int, bool[]> flaggedStates = null)
        {
            return new ScoringLogEntry
            {
                Id = Guid.NewGuid().ToString("N"),
                RangeLabel = rangeLabel ?? "",
                ScoredAt = DateTime.Now,
                GroupId = groupId,
                IsVariantMode = isVariantMode,
                VariantSetNo = isVariantMode && variantSetNo > 0 ? variantSetNo : 0,
                ProjectResults = CloneResults(projectResults),
                CompletedStates = CloneFlags(completedStates),
                FlaggedStates = CloneFlags(flaggedStates)
            };
        }

        public Dictionary<int, bool[]> GetCompletedStates()
        {
            return ToBoolArrays(CompletedStates);
        }

        public Dictionary<int, bool[]> GetFlaggedStates()
        {
            return ToBoolArrays(FlaggedStates);
        }

        static Dictionary<int, List<bool>> CloneResults(Dictionary<int, List<bool>> source)
        {
            var clone = new Dictionary<int, List<bool>>();
            if (source == null) return clone;
            foreach (var kv in source)
                clone[kv.Key] = kv.Value != null ? new List<bool>(kv.Value) : new List<bool>();
            return clone;
        }

        static Dictionary<int, List<bool>> CloneFlags(Dictionary<int, bool[]> source)
        {
            var clone = new Dictionary<int, List<bool>>();
            if (source == null) return clone;
            foreach (var kv in source)
                clone[kv.Key] = kv.Value != null ? new List<bool>(kv.Value) : new List<bool>();
            return clone;
        }

        static Dictionary<int, bool[]> ToBoolArrays(Dictionary<int, List<bool>> source)
        {
            var clone = new Dictionary<int, bool[]>();
            if (source == null) return clone;
            foreach (var kv in source)
                clone[kv.Key] = kv.Value != null ? kv.Value.ToArray() : Array.Empty<bool>();
            return clone;
        }
    }

    public static class ScoringLogStore
    {
        public const string SubjectExcel = "Excel";
        public const string SubjectWord = "Word";
        public const string SubjectPowerPoint = "PowerPoint";

        public static string GetPath(string subject)
        {
            return Path.Combine(ExamineeStore.DataDirectory, $"scoring-logs-{subject}.json");
        }

        public static List<ScoringLogEntry> Load(string subject)
        {
            try
            {
                string path = GetPath(subject);
                if (!File.Exists(path))
                    return new List<ScoringLogEntry>();

                var entries = JsonConvert.DeserializeObject<List<ScoringLogEntry>>(File.ReadAllText(path));
                if (entries == null)
                    return new List<ScoringLogEntry>();

                return entries
                    .OrderByDescending(e => e.ScoredAt)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ScoringLogStore] Load: " + ex.Message);
                return new List<ScoringLogEntry>();
            }
        }

        public static void Append(string subject, ScoringLogEntry entry)
        {
            if (entry == null) return;
            if (string.IsNullOrEmpty(entry.Id))
                entry.Id = Guid.NewGuid().ToString("N");
            if (entry.ScoredAt == default)
                entry.ScoredAt = DateTime.Now;

            var entries = Load(subject);
            entries.Insert(0, entry);
            Save(subject, entries);
        }

        /// <summary>大学情報の削除時に、Excel / Word / PowerPoint の採点ログをすべて消す。</summary>
        public static void ClearAll()
        {
            foreach (string subject in new[] { SubjectExcel, SubjectWord, SubjectPowerPoint })
            {
                try
                {
                    string path = GetPath(subject);
                    if (File.Exists(path))
                        File.Delete(path);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[ScoringLogStore] ClearAll: " + ex.Message);
                }
            }
        }

        static void Save(string subject, List<ScoringLogEntry> entries)
        {
            Directory.CreateDirectory(ExamineeStore.DataDirectory);
            File.WriteAllText(
                GetPath(subject),
                JsonConvert.SerializeObject(entries ?? new List<ScoringLogEntry>(), Formatting.Indented));
        }
    }
}
