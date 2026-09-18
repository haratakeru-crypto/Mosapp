namespace MosPracticeClient
{
    /// <summary>
    /// オフライン／送信失敗時QRの公開パス。Excel / Word / PowerPoint で同じ ID を使う。
    /// kouzakanri の MOS_PRACTICE_QR_PATH_ID（未設定時は同じ既定値）と一致させる。
    /// </summary>
    public static class MosPracticePublicQrPath
    {
        public const string DefaultPathId = "5cabe93ea53c64713796ac05fab9239f";

        public static string Normalize(string value)
        {
            string id = (value ?? "").Trim().Trim('/');
            return string.IsNullOrEmpty(id) ? DefaultPathId : id;
        }

        public static string BuildPagePath(string pathId)
        {
            return "/p/" + Normalize(pathId);
        }
    }
}
