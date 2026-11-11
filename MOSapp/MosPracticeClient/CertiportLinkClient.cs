using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MosPracticeClient
{
    public sealed class CertiportLinkResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
    }

    public static class CertiportLinkClient
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

        public static async Task<CertiportLinkResult> SubmitAsync(ExamineeProfile profile)
        {
            var result = new CertiportLinkResult();
            if (profile == null
                || string.IsNullOrWhiteSpace(profile.UniversityName)
                || string.IsNullOrWhiteSpace(profile.PersonName)
                || string.IsNullOrWhiteSpace(profile.CertiportDisplayName)
                || string.IsNullOrWhiteSpace(profile.CertiportCorrectName))
            {
                result.Error = "大学名・氏名と認定証の名前がそろっていません";
                return result;
            }

            var settings = PracticeSubmitConfig.Load();
            if (string.IsNullOrEmpty(settings.BaseUrl) || string.IsNullOrEmpty(settings.IngestKey))
            {
                result.Error = "送信先が設定されていません";
                return result;
            }

            var body = new JObject
            {
                ["clientId"] = PresenceClient.GetOrCreateClientId(),
                ["universityName"] = profile.UniversityName ?? "",
                ["personName"] = profile.PersonName ?? "",
                ["certiportLoginEmail"] = profile.CertiportLoginEmail ?? "",
                ["certiportDisplayName"] = profile.CertiportDisplayName ?? "",
                ["correctName"] = profile.CertiportCorrectName ?? "",
                ["namesMatch"] = profile.CertiportNameMatchesMos
            };
            if (!profile.CertiportNameMatchesMos)
            {
                string incorrect = CertiportPortal.NamesMatch(profile.CertiportCorrectName, profile.PersonName)
                    ? profile.CertiportDisplayName
                    : profile.PersonName;
                body["incorrectName"] = incorrect ?? "";
            }
            if (!string.IsNullOrWhiteSpace(profile.ClassroomName))
                body["classroomName"] = profile.ClassroomName;
            if (!string.IsNullOrWhiteSpace(profile.CertiportVerifiedAt))
                body["verifiedAt"] = profile.CertiportVerifiedAt;

            try
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, settings.BaseUrl + "/api/mos-practice/certiport-link"))
                {
                    request.Headers.TryAddWithoutValidation("X-MOS-Practice-Key", settings.IngestKey);
                    request.Content = new StringContent(body.ToString(Formatting.None), Encoding.UTF8, "application/json");
                    using (var response = await Http.SendAsync(request).ConfigureAwait(false))
                    {
                        string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        if (response.IsSuccessStatusCode)
                        {
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
