using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace MosPracticeClient
{
    public static class ExamineeStore
    {
        public static string DataDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MOSapp");
            }
        }

        public static string ExamineePath => Path.Combine(DataDirectory, "examinee.json");
        public static string LookupsPath => Path.Combine(DataDirectory, "lookups.json");

        public static bool IsRegistered
        {
            get
            {
                var profile = Load();
                return profile != null
                    && !string.IsNullOrWhiteSpace(profile.UniversityName)
                    && !string.IsNullOrWhiteSpace(profile.PersonName);
            }
        }

        public static ExamineeProfile Load()
        {
            try
            {
                if (!File.Exists(ExamineePath)) return null;
                var profile = JsonConvert.DeserializeObject<ExamineeProfile>(File.ReadAllText(ExamineePath));
                if (profile == null) return null;
                if (profile.SubmittedSubjects == null)
                    profile.SubmittedSubjects = new List<string>();
                return profile;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExamineeStore] Load: " + ex.Message);
                return null;
            }
        }

        public static void Save(ExamineeProfile profile)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            Directory.CreateDirectory(DataDirectory);
            if (profile.SubmittedSubjects == null)
                profile.SubmittedSubjects = new List<string>();
            File.WriteAllText(ExamineePath, JsonConvert.SerializeObject(profile, Formatting.Indented));
            SessionWatchLauncher.EnsureRunning();
        }

        public static void Delete()
        {
            try
            {
                if (File.Exists(ExamineePath))
                    File.Delete(ExamineePath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExamineeStore] Delete: " + ex.Message);
            }
            SessionWatchLauncher.Stop();
        }

        public static bool HasSubmitted(string subject)
        {
            var profile = Load();
            if (profile?.SubmittedSubjects == null || string.IsNullOrEmpty(subject)) return false;
            return profile.SubmittedSubjects.Exists(s =>
                string.Equals(s, subject, StringComparison.OrdinalIgnoreCase));
        }

        public static void MarkSubmitted(string subject)
        {
            var profile = Load() ?? new ExamineeProfile();
            if (profile.SubmittedSubjects == null)
                profile.SubmittedSubjects = new List<string>();
            if (!HasSubmitted(subject) && !string.IsNullOrEmpty(subject))
                profile.SubmittedSubjects.Add(subject);
            Save(profile);
        }

        public static LookupsPayload LoadLookupsCache()
        {
            try
            {
                if (!File.Exists(LookupsPath)) return new LookupsPayload();
                return JsonConvert.DeserializeObject<LookupsPayload>(File.ReadAllText(LookupsPath))
                    ?? new LookupsPayload();
            }
            catch
            {
                return new LookupsPayload();
            }
        }

        public static void SaveLookupsCache(LookupsPayload payload)
        {
            if (payload == null) return;
            Directory.CreateDirectory(DataDirectory);
            File.WriteAllText(LookupsPath, JsonConvert.SerializeObject(payload, Formatting.Indented));
        }
    }
}
