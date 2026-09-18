using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MosPracticeClient
{
    public static class PracticeSubmitConfig
    {
        public static PracticeSubmitSettings Load()
        {
            var settings = new PracticeSubmitSettings();
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (!File.Exists(path)) return settings;
                var root = JObject.Parse(File.ReadAllText(path));
                var node = root["practiceSubmit"] as JObject;
                if (node == null) return settings;
                settings.BaseUrl = (node.Value<string>("baseUrl") ?? "").Trim().TrimEnd('/');
                settings.IngestKey = (node.Value<string>("ingestKey") ?? "").Trim();
                settings.QrPathId = MosPracticePublicQrPath.Normalize(node.Value<string>("qrPathId"));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PracticeSubmitConfig] " + ex.Message);
            }
            return settings;
        }
    }
}
