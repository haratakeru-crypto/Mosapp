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

        [JsonProperty("projectResults")]
        public Dictionary<int, List<bool>> ProjectResults { get; set; }

        [JsonIgnore]
        public string DisplayTitle
        {
            get
            {
                string label = string.IsNullOrWhiteSpace(RangeLabel) ? "採点" : RangeLabel.Trim();
                return $"{label} {ScoredAt:M/d H:mm}";
            }
        }

        public static ScoringLogEntry Create(string rangeLabel, int groupId, Dictionary<int, List<bool>> projectResults)
        {
            return new ScoringLogEntry
            {
                Id = Guid.NewGuid().ToString("N"),
                RangeLabel = rangeLabel ?? "",
                ScoredAt = DateTime.Now,
                GroupId = groupId,
                ProjectResults = CloneResults(projectResults)
            };
        }

        static Dictionary<int, List<bool>> CloneResults(Dictionary<int, List<bool>> source)
        {
            var clone = new Dictionary<int, List<bool>>();
            if (source == null) return clone;
            foreach (var kv in source)
                clone[kv.Key] = kv.Value != null ? new List<bool>(kv.Value) : new List<bool>();
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

        static void Save(string subject, List<ScoringLogEntry> entries)
        {
            Directory.CreateDirectory(ExamineeStore.DataDirectory);
            File.WriteAllText(
                GetPath(subject),
                JsonConvert.SerializeObject(entries ?? new List<ScoringLogEntry>(), Formatting.Indented));
        }
    }
}
