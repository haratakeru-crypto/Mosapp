using System;

using System.Diagnostics;

using System.IO;

using System.Runtime.InteropServices;

using System.Threading;

using System.Threading.Tasks;

using WordApp = Microsoft.Office.Interop.Word.Application;



namespace Libraries

{

    /// <summary>

    /// Word を通常起動し、VSTO アドインのハートビートが更新されるまで待機する。

    /// </summary>

    public static class WordApplicationManager

    {

        private const int DefaultAcquireTimeoutMs = 10000;

        private const int ExistingWordHeartbeatWaitMs = 1500;

        private const int VstoHeartbeatWaitAfterLaunchMs = 8000;

        private const int WordQuitWaitMs = 5000;

        private const int HeartbeatPollIntervalMs = 100;



        private static int _warmupStarted;



        [DllImport("user32.dll")]

        private static extern bool AllowSetForegroundWindow(uint dwProcessId);



        /// <summary>

        /// 起動直後にバックグラウンドで Word + VSTO をウォームアップする（プロジェクト開始時の待ちを短縮）。

        /// </summary>

        public static void StartBackgroundWordWarmup()

        {

            if (Interlocked.CompareExchange(ref _warmupStarted, 1, 0) != 0)

                return;



            Task.Run(() =>

            {

                try

                {

                    if (LogReader.IsVstoHeartbeatFresh(300))

                    {

                        System.Diagnostics.Debug.WriteLine("[WordApplicationManager] Warmup skipped: heartbeat already fresh.");

                        return;

                    }



                    VSTOInstallerHelper.EnsureAddInReadyForExam(out string issue);

                    if (!string.IsNullOrEmpty(issue))

                        System.Diagnostics.Debug.WriteLine($"[WordApplicationManager] Warmup VSTO prep: {issue}");



                    AcquireWordApplicationForExam(makeVisible: false, timeoutMs: DefaultAcquireTimeoutMs);

                    System.Diagnostics.Debug.WriteLine("[WordApplicationManager] Background warmup completed.");

                }

                catch (Exception ex)

                {

                    System.Diagnostics.Debug.WriteLine($"[WordApplicationManager] Background warmup failed: {ex.Message}");

                }

            });

        }



        /// <summary>

        /// 試験用: Release VSTO を有効化し、ハートビートが取れる Word を返す（既存 Word に VSTO が無ければ再起動）。

        /// </summary>

        public static WordApp AcquireWordApplicationForExam(bool makeVisible = true, int timeoutMs = DefaultAcquireTimeoutMs)

        {

            VSTOInstallerHelper.EnsureAddInReadyForExam(out _);



            WordApp existing = TryGetActiveWordApplication();

            if (existing != null)

            {

                TrySetVisible(existing, makeVisible);

                if (WaitForVstoHeartbeat(ExistingWordHeartbeatWaitMs))

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

                Thread.Sleep(HeartbeatPollIntervalMs);

            }

            return LogReader.IsVstoHeartbeatFresh(300);

        }



        private static WordApp LaunchWordAndWaitForVsto(bool makeVisible, int timeoutMs)

        {

            int launchedPid = StartWordProcess();

            if (launchedPid <= 0)

                throw new InvalidOperationException("Word を起動できませんでした。");



            WordApp app = WaitForActiveWordApplication(timeoutMs);

            if (app == null)

                throw new InvalidOperationException("起動後の Word へ接続できませんでした。");



            TrySetVisible(app, makeVisible);

            int vstoWaitMs = Math.Min(timeoutMs, VstoHeartbeatWaitAfterLaunchMs);

            if (!WaitForVstoHeartbeat(vstoWaitMs))

                System.Diagnostics.Debug.WriteLine("[WordApplicationManager] VSTO heartbeat not detected after Word launch.");



            return app;

        }



        private static void TryQuitWordAndWait(WordApp app)

        {

            try { app.Quit(SaveChanges: false); } catch { /* ignore */ }



            var sw = Stopwatch.StartNew();

            while (sw.ElapsedMilliseconds < WordQuitWaitMs)

            {

                if (Process.GetProcessesByName("WINWORD").Length == 0)

                    return;

                Thread.Sleep(200);

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

                Thread.Sleep(HeartbeatPollIntervalMs);

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


