using System.Diagnostics;

namespace Libraries
{
    /// <summary>
    /// 採点処理のパフォーマンス計測（Debug 出力）。[Perf] プレフィックスでフィルタ可能。
    /// </summary>
    public static class PPGradingPerf
    {
        /// <summary>
        /// false にすると [Perf] ログを出さない。
        /// </summary>
        public static bool Enabled { get; set; } = true;

        public static void Log(string category, long elapsedMs, string detail = null)
        {
            if (!Enabled)
                return;
            string suffix = string.IsNullOrEmpty(detail) ? string.Empty : " " + detail;
            Debug.WriteLine($"[Perf] {category}: {elapsedMs} ms{suffix}");
        }
    }
}
