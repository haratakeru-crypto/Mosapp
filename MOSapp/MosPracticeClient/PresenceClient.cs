using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MosPracticeClient
{
    public static class PresenceClient
    {
        static readonly HttpClient Http = CreateClient();

        public static string ClientIdPath => Path.Combine(ExamineeStore.DataDirectory, "client-id.txt");
        public static string PresenceBaseUrlPath => Path.Combine(ExamineeStore.DataDirectory, "presence-baseurl.txt");
        public static string PresenceKeyPath => Path.Combine(ExamineeStore.DataDirectory, "presence-key.txt");

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

        public static string GetOrCreateClientId()
        {
            try
            {
                Directory.CreateDirectory(ExamineeStore.DataDirectory);
                if (File.Exists(ClientIdPath))
                {
                    string existing = (File.ReadAllText(ClientIdPath) ?? "").Trim();
                    if (existing.Length >= 8) return existing;
                }
                string created = Guid.NewGuid().ToString("N");
                File.WriteAllText(ClientIdPath, created);
                return created;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PresenceClient] clientId: " + ex.Message);
                return Guid.NewGuid().ToString("N");
            }
        }

        public static void NotifyRegistered(ExamineeProfile profile)
        {
            WriteEndpointFiles();
            Task.Run(() => PostAsync("register", profile));
        }

        public static void NotifyUnregisteredAndWait()
        {
            try
            {
                PostAsync("unregister", null).Wait(TimeSpan.FromSeconds(3));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PresenceClient] unregister: " + ex.Message);
            }
        }

        static void WriteEndpointFiles()
        {
            try
            {
                var settings = PracticeSubmitConfig.Load();
                Directory.CreateDirectory(ExamineeStore.DataDirectory);
                File.WriteAllText(PresenceBaseUrlPath, settings.BaseUrl ?? "");
                File.WriteAllText(PresenceKeyPath, settings.IngestKey ?? "");
                GetOrCreateClientId();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PresenceClient] endpoint: " + ex.Message);
            }
        }

        static async Task PostAsync(string action, ExamineeProfile profile)
        {
            var settings = PracticeSubmitConfig.Load();
            if (string.IsNullOrEmpty(settings.BaseUrl) || string.IsNullOrEmpty(settings.IngestKey))
                return;

            var body = new JObject
            {
                ["clientId"] = GetOrCreateClientId(),
                ["action"] = action ?? ""
            };
            if (profile != null)
            {
                body["universityName"] = profile.UniversityName ?? "";
                body["personName"] = profile.PersonName ?? "";
                if (!string.IsNullOrWhiteSpace(profile.ClassroomName))
                    body["classroomName"] = profile.ClassroomName;
            }

            using (var request = new HttpRequestMessage(HttpMethod.Post, settings.BaseUrl + "/api/mos-practice/presence"))
            {
                request.Headers.TryAddWithoutValidation("X-MOS-Practice-Key", settings.IngestKey);
                request.Content = new StringContent(body.ToString(Formatting.None), Encoding.UTF8, "application/json");
                using (var response = await Http.SendAsync(request).ConfigureAwait(false))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        System.Diagnostics.Debug.WriteLine("[PresenceClient] HTTP " + (int)response.StatusCode + " " + json);
                    }
                }
            }
        }
    }
}
