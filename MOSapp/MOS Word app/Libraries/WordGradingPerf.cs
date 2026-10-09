using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Libraries
{
    /// <summary>
    /// 採点処理のパフォーマンス計測。処理中はメモリへ蓄積し、
    /// 一括採点終了時に %TEMP%\mos_word_grading_perf.log を直近1回の内容で置き換える。
    /// </summary>
    public static class WordGradingPerf
    {
        private static readonly object Sync = new object();
        private static readonly Dictionary<string, Sample> Samples = new Dictionary<string, Sample>(StringComparer.Ordinal);
        private static readonly List<string> Details = new List<string>();
        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "mos_word_grading_perf.log");

        /// <summary>
        /// false にすると計測を蓄積しない。
        /// </summary>
        public static bool Enabled { get; set; } = true;

        public static void BeginSession(string name)
        {
            if (!Enabled)
                return;

            lock (Sync)
            {
                Samples.Clear();
                Details.Clear();
                Details.Add($"===== {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {name} =====");
            }
        }

        public static void Log(string category, long elapsedMs, string detail = null)
        {
            if (!Enabled)
                return;

            lock (Sync)
            {
                if (!Samples.TryGetValue(category, out Sample sample))
                {
                    sample = new Sample();
                    Samples[category] = sample;
                }
                sample.Add(elapsedMs);
                if (!string.IsNullOrWhiteSpace(detail))
                    Details.Add($"{category}: {elapsedMs} ms {detail}");
            }
        }

        public static void EndSession()
        {
            if (!Enabled)
                return;

            string report;
            lock (Sync)
            {
                if (Details.Count == 0)
                    return;

                var builder = new StringBuilder();
                foreach (string detail in Details)
                    builder.AppendLine(detail);

                builder.AppendLine("--- summary ---");
                foreach (var pair in Samples.OrderBy(p => p.Key, StringComparer.Ordinal))
                {
                    Sample sample = pair.Value;
                    builder.AppendLine(
                        $"{pair.Key}: count={sample.Count} total={sample.TotalMs} ms avg={sample.AverageMs} ms max={sample.MaxMs} ms");
                }
                builder.AppendLine();
                report = builder.ToString();
                Samples.Clear();
                Details.Clear();
            }

            try
            {
                File.WriteAllText(LogPath, report, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[WordGradingPerf] " + ex.Message);
            }
        }

        private sealed class Sample
        {
            public int Count { get; private set; }
            public long TotalMs { get; private set; }
            public long MaxMs { get; private set; }
            public long AverageMs => Count == 0 ? 0 : TotalMs / Count;

            public void Add(long elapsedMs)
            {
                Count++;
                TotalMs += elapsedMs;
                if (elapsedMs > MaxMs)
                    MaxMs = elapsedMs;
            }
        }
    }
}
