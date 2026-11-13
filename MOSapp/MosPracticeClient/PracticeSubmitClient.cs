using System;
using System.Globalization;
using System.Linq;
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

        /// <summary>
        /// 座席表へ初回送信。類題は range に点数と採点範囲（例: 類題:700 P1～3）を載せ、×は wrongCount に任せる。
        /// </summary>
        public static async Task<PracticeSubmitResult> SubmitInitialAsync(
            int wrongCount,
            string subject,
            bool isVariant = false,
            int? practiceTotalScore = null)
        {
            var settings = PracticeSubmitConfig.Load();
            var profile = ExamineeStore.Load();
            var scoredAt = DateTime.Now;
            string attempt = "";
            string range = "";
            if (!string.IsNullOrEmpty(subject))
            {
                var logs = ScoringLogStore.Load(subject);
                if (logs.Count > 0)
                {
                    var latest = logs.FirstOrDefault(e => e.IsVariantMode == isVariant) ?? logs[0];
                    if (latest.ScoredAt != default(DateTime))
                        scoredAt = latest.ScoredAt;
                    if (!isVariant)
                    {
                        attempt = logs.Count.ToString(CultureInfo.InvariantCulture);
                        range = ScoringRangeSelection.ShortenRangeLabel(latest.RangeLabel);
                    }
                    else
                    {
                        range = BuildVariantRangeLabel(practiceTotalScore, latest.RangeLabel);
                    }
                }
            }
            if (isVariant && string.IsNullOrEmpty(range))
                range = BuildVariantRangeLabel(practiceTotalScore, null);
            string at = scoredAt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);

            var result = new PracticeSubmitResult();
            if (profile == null
                || string.IsNullOrWhiteSpace(profile.UniversityName)
                || string.IsNullOrWhiteSpace(profile.PersonName))
            {
                result.Error = "大学名・氏名が登録されていません";
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
                settings.QrPathId,
                attempt,
                range);

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
                if (!string.IsNullOrEmpty(attempt))
                    body["attempt"] = attempt;
                if (!string.IsNullOrEmpty(range))
                    body["range"] = range;
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

        /// <summary>類題用 range。点数があれば「類題:700」、範囲があれば続けて「 P1～3」。</summary>
        internal static string BuildVariantRangeLabel(int? practiceTotalScore, string rangeLabel)
        {
            string shortRange = ScoringRangeSelection.ShortenRangeLabel(rangeLabel);
            if (practiceTotalScore.HasValue)
            {
                string withScore = "類題:" + practiceTotalScore.Value.ToString(CultureInfo.InvariantCulture);
                return shortRange.Length > 0 ? withScore + " " + shortRange : withScore;
            }

            return shortRange.Length > 0 ? "類題 " + shortRange : "類題";
        }
    }
}
