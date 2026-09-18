using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MosPracticeClient
{
    public static class LookupsClient
    {
        static readonly HttpClient Http = CreateClient();

        static string BundledPath =>
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "universities.json");

        static HttpClient CreateClient()
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
            }
            catch { }
            var client = new HttpClient();
            client.Timeout = TimeSpan.FromSeconds(8);
            return client;
        }

        public static List<UniversityLookup> GetCachedOrEmpty()
        {
            var cache = ExamineeStore.LoadLookupsCache();
            if (HasUniversities(cache))
                return Normalize(cache.Universities);

            return LoadBundledUniversities();
        }

        public static async Task<List<UniversityLookup>> RefreshAsync()
        {
            var local = GetCachedOrEmpty();
            var settings = PracticeSubmitConfig.Load();
            if (string.IsNullOrEmpty(settings.BaseUrl)) return local;

            try
            {
                string json = await Http.GetStringAsync(settings.BaseUrl + "/api/mos-practice/lookups").ConfigureAwait(false);
                var payload = JsonConvert.DeserializeObject<LookupsPayload>(json) ?? new LookupsPayload();
                var universities = Normalize(payload.Universities);
                payload.Universities = universities;
                ExamineeStore.SaveLookupsCache(payload);
                return universities;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[LookupsClient] " + ex.Message);
                return local;
            }
        }

        public static IEnumerable<UniversityLookup> FilterUniversities(IEnumerable<UniversityLookup> source, string query)
        {
            var list = source ?? Enumerable.Empty<UniversityLookup>();
            if (string.IsNullOrWhiteSpace(query)) return Enumerable.Empty<UniversityLookup>();
            string q = query.Trim();
            return list.Where(u => (u.Name ?? "").IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        static List<UniversityLookup> LoadBundledUniversities()
        {
            try
            {
                if (!File.Exists(BundledPath)) return new List<UniversityLookup>();
                var payload = JsonConvert.DeserializeObject<LookupsPayload>(File.ReadAllText(BundledPath));
                return Normalize(payload?.Universities);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[LookupsClient] bundled: " + ex.Message);
                return new List<UniversityLookup>();
            }
        }

        static bool HasUniversities(LookupsPayload payload)
        {
            return payload?.Universities != null && payload.Universities.Count > 0;
        }

        static List<UniversityLookup> Normalize(List<UniversityLookup> source)
        {
            var list = source ?? new List<UniversityLookup>();
            foreach (var item in list)
            {
                if (item.Classrooms == null) item.Classrooms = new List<string>();
            }
            return list;
        }
    }
}
