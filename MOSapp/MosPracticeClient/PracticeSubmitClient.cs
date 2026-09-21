using System;
using System.Globalization;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MosPracticeClient
{
    public sealed class PracticeSubmitResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string QrUrl { get; set; }
    }

    public static class PracticeSubmitClient
    {
        static readonly HttpClient Http = CreateClient();

        static HttpClient CreateClient()
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
            }
            catch { }
            var client = new HttpClient();
            client.Timeout = TimeSpan.FromSeconds(20);
            return client;
        }

        public static async Task<PracticeSubmitResult> SubmitInitialAsync(int wrongCount, string subject)
        {
            var settings = PracticeSubmitConfig.Load();
            var profile = ExamineeStore.Load();
            var scoredAt = DateTime.Now;
            string at = scoredAt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);

            var result = new PracticeSubmitResult();
            if (profile == null
                || string.IsNullOrWhiteSpace(profile.UniversityName)
                || string.IsNullOrWhiteSpace(profile.PersonName))
            {
                result.Error = "大学名・氏名が登録されていません";
                return result;
            }

            if (ExamineeStore.HasSubmitted(subject))
            {
                result.Success = true;
                return result;
            }

            if (string.IsNullOrEmpty(settings.BaseUrl) || string.IsNullOrEmpty(settings.IngestKey))
            {
                result.Error = "送信先が設定されていません";
                return result;
            }

            result.QrUrl = QrUrlBuilder.BuildSubmitPageUrl(
                settings.BaseUrl,
                settings.IngestKey,
                profile.UniversityName,
                profile.PersonName,
                profile.ClassroomName,
                wrongCount,
                scoredAt,
                subject,
                settings.QrPathId);

            try
            {
                var body = new JObject
                {
                    ["universityName"] = profile.UniversityName,
                    ["personName"] = profile.PersonName,
                    ["wrongCount"] = wrongCount,
                    ["scoredAt"] = at,
                    ["subject"] = subject ?? ""
                };
                if (!string.IsNullOrWhiteSpace(profile.ClassroomName))
                    body["classroomName"] = profile.ClassroomName;

                using (var request = new HttpRequestMessage(HttpMethod.Post, settings.BaseUrl + "/api/mos-practice/submit"))
                {
                    request.Headers.TryAddWithoutValidation("X-MOS-Practice-Key", settings.IngestKey);
                    request.Content = new StringContent(body.ToString(Formatting.None), Encoding.UTF8, "application/json");
                    using (var response = await Http.SendAsync(request).ConfigureAwait(false))
                    {
                        string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        if (response.IsSuccessStatusCode)
                        {
                            ExamineeStore.MarkSubmitted(subject);
                            result.Success = true;
                            return result;
                        }

                        try
                        {
                            var parsed = JObject.Parse(json);
                            result.Error = parsed.Value<string>("error") ?? ("HTTP " + (int)response.StatusCode);
                        }
                        catch
                        {
                            result.Error = "HTTP " + (int)response.StatusCode;
                        }
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                result.Error = ex.Message;
                return result;
            }
        }
    }
}
