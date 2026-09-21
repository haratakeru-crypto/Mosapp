using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace MosPracticeClient
{
    public static class SessionWatchLauncher
    {
        public const string MutexName = "Local\\MosPracticeSessionWatch";
        const string ExeName = "MosPracticeSessionWatch.exe";

        static string ExePath
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ExeName); }
        }

        public static void EnsureRunning()
        {
            try
            {
                if (!File.Exists(ExePath))
                {
                    Debug.WriteLine("[SessionWatchLauncher] exe missing: " + ExePath);
                    return;
                }

                var existing = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(ExeName));
                if (existing.Length > 0) return;

                var start = new ProcessStartInfo
                {
                    FileName = ExePath,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                Process.Start(start);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SessionWatchLauncher] " + ex.Message);
            }
        }

        public static void Stop()
        {
            try
            {
                foreach (var proc in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(ExeName)))
                {
                    try { proc.Kill(); }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[SessionWatchLauncher] Stop: " + ex.Message);
            }
        }
    }
}
