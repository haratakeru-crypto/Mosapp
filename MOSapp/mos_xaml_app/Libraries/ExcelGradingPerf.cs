using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Libraries
{
    /// <summary>
    /// Excel一括採点のパフォーマンス計測。処理中はメモリへ蓄積し、
    /// 一括採点の記録は %TEMP%\mos_excel_grading_perf.log に直近1回だけ残す。
    /// </summary>
    public static class ExcelGradingPerf
    {
        private static readonly object Sync = new object();
        private static readonly Dictionary<string, Sample> Samples = new Dictionary<string, Sample>(StringComparer.Ordinal);
        private static readonly List<string> Details = new List<string>();
        private static readonly List<string> Preamble = new List<string>();
        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "mos_excel_grading_perf.log");
        private static bool _sessionOpen;

        public static bool Enabled { get; set; } = true;

        public static void BeginSession(string name)
        {
            if (!Enabled)
                return;

            string text;
            lock (Sync)
            {
                Samples.Clear();
                var carried = new List<string>(Preamble);
                Preamble.Clear();
                Details.Clear();
                Details.Add($"===== {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {name} =====");
                Details.AddRange(carried);
                _sessionOpen = true;
                text = string.Join(Environment.NewLine, Details) + Environment.NewLine;
            }

            try
            {
                ReplaceFile(text);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ExcelGradingPerf] " + ex.Message);
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

        /// <summary>
        /// 採点セッション開始前のイベントを、まとめ出力を待たずに記録する。
        /// </summary>
        public static void LogImmediate(string category, long elapsedMs, string detail = null)
        {
            if (!Enabled)
                return;

            try
            {
                string line = string.IsNullOrWhiteSpace(detail)
                    ? $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {category}: {elapsedMs} ms"
                    : $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {category}: {elapsedMs} ms {detail}";
                string text;
                bool replace;
                lock (Sync)
                {
                    if (_sessionOpen)
                    {
                        Details.Add(line);
                        replace = false;
                        text = line + Environment.NewLine;
                    }
                    else
                    {
                        Preamble.Add(line);
                        replace = true;
                        text = string.Join(Environment.NewLine, Preamble) + Environment.NewLine;
                    }
                }

                if (replace)
                    ReplaceFile(text);
                else
                    File.AppendAllText(LogPath, text, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ExcelGradingPerf] " + ex.Message);
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
                _sessionOpen = false;
            }

            try
            {
                File.WriteAllText(LogPath, report, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ExcelGradingPerf] " + ex.Message);
            }
        }

        private static void ReplaceFile(string text)
        {
            File.WriteAllText(LogPath, text, new UTF8Encoding(false));
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
