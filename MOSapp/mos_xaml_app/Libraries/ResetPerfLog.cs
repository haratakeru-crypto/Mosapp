using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace Libraries
{
    /// <summary>
    /// アプリバーのリセット所要時間。%TEMP%\mos_reset_perf.log は直近1回だけ残す。
    /// </summary>
    public static class ResetPerfLog
    {
        static readonly object Sync = new object();
        static readonly string LogPath = Path.Combine(Path.GetTempPath(), "mos_reset_perf.log");

        public static void Begin(string app, int projectId)
        {
            WriteLine($"===== {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {app} project={projectId} reset =====", replace: true);
        }

        public static void Write(string app, int projectId, string stage, long elapsedMs, string path, string detail = null)
        {
            string line = string.IsNullOrWhiteSpace(detail)
                ? $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {app} project={projectId} {stage} {elapsedMs} ms path={path}"
                : $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {app} project={projectId} {stage} {elapsedMs} ms path={path} {detail}";
            WriteLine(line, replace: false);
        }

        static void WriteLine(string line, bool replace)
        {
            try
            {
                lock (Sync)
                {
                    string text = line + Environment.NewLine;
                    if (replace)
                        File.WriteAllText(LogPath, text, new UTF8Encoding(false));
                    else
                        File.AppendAllText(LogPath, text, new UTF8Encoding(false));
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ResetPerfLog] " + ex.Message);
            }
        }
    }
}
