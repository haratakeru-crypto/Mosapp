using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using WordApp = Microsoft.Office.Interop.Word.Application;

namespace Libraries
{
    /// <summary>
    /// Word を通常起動し、VSTO アドインのハートビートが更新されるまで待機する。
    /// </summary>
    public static class WordApplicationManager
    {
        [DllImport("user32.dll")]
        private static extern bool AllowSetForegroundWindow(uint dwProcessId);

        /// <summary>
        /// 試験用: Release VSTO を有効化し、ハートビートが取れる Word を返す（既存 Word に VSTO が無ければ再起動）。
        /// </summary>
        public static WordApp AcquireWordApplicationForExam(bool makeVisible = true, int timeoutMs = 30000)
        {
            VSTOInstallerHelper.EnsureAddInReadyForExam(out _);

            WordApp existing = TryGetActiveWordApplication();
            if (existing != null)
            {
                TrySetVisible(existing, makeVisible);
                if (WaitForVstoHeartbeat(4000))
                    return existing;

                TryQuitWordAndWait(existing);
                try { Marshal.FinalReleaseComObject(existing); } catch { }
            }

            return LaunchWordAndWaitForVsto(makeVisible, timeoutMs);
        }

        public static bool WaitForVstoHeartbeat(int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (LogReader.IsVstoHeartbeatFresh(300))
                    return true;
                Thread.Sleep(250);
            }
            return LogReader.IsVstoHeartbeatFresh(300);
        }

        private static WordApp LaunchWordAndWaitForVsto(bool makeVisible, int timeoutMs)
        {
            ClearVstoHeartbeat();
            int launchedPid = StartWordProcess();
            if (launchedPid <= 0)
                throw new InvalidOperationException("Word を起動できませんでした。");

            WordApp app = WaitForActiveWordApplication(timeoutMs);
            if (app == null)
                throw new InvalidOperationException("起動後の Word へ接続できませんでした。");

            TrySetVisible(app, makeVisible);
            if (!WaitForVstoHeartbeat(Math.Min(timeoutMs, 25000)))
                System.Diagnostics.Debug.WriteLine("[WordApplicationManager] VSTO heartbeat not detected after Word launch.");

            return app;
        }

        private static void TryQuitWordAndWait(WordApp app)
        {
            try { app.Quit(SaveChanges: false); } catch { /* ignore */ }

            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < 15000)
            {
                if (Process.GetProcessesByName("WINWORD").Length == 0)
                    return;
                Thread.Sleep(300);
            }

            foreach (Process p in Process.GetProcessesByName("WINWORD"))
            {
                try
                {
                    if (!p.HasExited)
                        p.Kill();
                }
                catch { /* ignore */ }
                finally
                {
                    try { p.Dispose(); } catch { }
                }
            }
        }

        private static void ClearVstoHeartbeat()
        {
            try
            {
                string path = LogReader.GetVstoHeartbeatPath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch { /* ignore */ }
        }

        private static WordApp TryGetActiveWordApplication()
        {
            try
            {
                return (WordApp)Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                return null;
            }
        }

        private static WordApp WaitForActiveWordApplication(int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                var app = TryGetActiveWordApplication();
                if (app != null)
                    return app;
                Thread.Sleep(250);
            }
            return null;
        }

        private static void TrySetVisible(WordApp app, bool makeVisible)
        {
            if (app == null) return;
            try { app.Visible = makeVisible; } catch { }
        }

        private static int StartWordProcess()
        {
            string[] candidates =
            {
                "winword.exe",
                @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
            };

            foreach (string path in candidates)
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    };
                    Process proc = Process.Start(psi);
                    if (proc == null)
                        continue;

                    try { AllowSetForegroundWindow((uint)proc.Id); } catch { }
                    return proc.Id;
                }
                catch
                {
                    // try next
                }
            }

            return -1;
        }
    }
}
