using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace Libraries
{
    /// <summary>
    /// Excel を「COMオートメーション起動」ではなく通常起動に寄せて取得するためのヘルパー。
    /// VSTO アドインがロードされるまで（Startup completed）待機してから返す。
    /// </summary>
    public static class ExcelApplicationManager
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AllowSetForegroundWindow(uint dwProcessId);

        private static readonly string AddinDiagPath = ExcelLogReader.GetDiagnosticLogPath();
        private const string AddinStartupCompletedMarker = "Startup completed";

        public static ExcelApp GetOrCreateExcelApplication(bool makeVisible, int timeoutMs = 30000)
        {
            ExcelApp app = null;
            try
            {
                // 既に起動中ならそれを使う（アドインロード済み前提で、PIDが取れれば待機も行う）
                app = TryGetActiveExcelApplication();
                if (app != null)
                {
                    TrySetVisible(app, makeVisible);
                    WaitForVstoStartupIfPossible(app, timeoutMs);
                    return app;
                }
            }
            catch
            {
                // 次の通常起動にフォールバック
            }

            // Excel が見つからない場合は「通常起動」する（COM起動の new ExcelApp() を避ける）
            int launchedPid = -1;
            bool launchSucceeded = false;
            try
            {
                var swLaunch = Stopwatch.StartNew();
                launchedPid = StartExcelProcess();
                if (launchedPid <= 0)
                    throw new InvalidOperationException("Excel を起動できませんでした。");

                app = WaitForActiveExcelApplicationForPid(launchedPid, timeoutMs);
                if (app == null)
                {
                    int remaining = Math.Max(2000, timeoutMs - (int)swLaunch.ElapsedMilliseconds);
                    app = WaitForActiveExcelApplication(remaining);
                }

                if (app == null)
                    throw new InvalidOperationException("起動後の Excel へ接続できませんでした。");

                TrySetVisible(app, makeVisible);
                WaitForVstoStartupIfPossible(app, 5000);
                launchSucceeded = true;
                return app;
            }
            finally
            {
                if (!launchSucceeded && launchedPid > 0)
                {
                    EnsureExcelProcessExited(
                        launchedPid,
                        10000,
                        5000,
                        "[GetOrCreateExcelApplication] failed launch cleanup");
                }
            }
        }

        /// <summary>
        /// 起動済みの Excel のみ ROT から取得する。新規プロセスは起動しない（シェル起動後の COM 接続用）。
        /// </summary>
        public static ExcelApp TryAttachRunningExcelApplication(bool makeVisible, int timeoutMs = 15000)
        {
            var app = WaitForActiveExcelApplication(timeoutMs);
            if (app == null)
                return null;

            TrySetVisible(app, makeVisible);
            int vstoWaitMs = Math.Min(timeoutMs, 8000);
            WaitForVstoStartupIfPossible(app, vstoWaitMs);
            return app;
        }

        private static ExcelApp TryGetActiveExcelApplication()
        {
            try
            {
                return (ExcelApp)Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                return null;
            }
        }

        private static ExcelApp WaitForActiveExcelApplication(int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                var app = TryGetActiveExcelApplication();
                if (app != null) return app;
                Thread.Sleep(250);
            }
            return null;
        }

        /// <summary>
        /// 指定 PID の Excel プロセスに対応する Application を ROT から取得できるまで待つ。
        /// Hwnd がまだ取れない起動直後は PID が -1 になり得るため、その間は再試行する。
        /// </summary>
        private static ExcelApp WaitForActiveExcelApplicationForPid(int expectedPid, int timeoutMs)
        {
            if (expectedPid <= 0)
                return WaitForActiveExcelApplication(timeoutMs);

            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                var app = TryGetActiveExcelApplication();
                if (app != null)
                {
                    int pid = TryGetExcelProcessId(app);
                    if (pid == expectedPid)
                        return app;
                }
                Thread.Sleep(250);
            }

            return null;
        }

        private static void TrySetVisible(ExcelApp app, bool makeVisible)
        {
            if (app == null) return;
            try { app.Visible = makeVisible; } catch { }
        }

        private static void WaitForVstoStartupIfPossible(ExcelApp app, int timeoutMs)
        {
            int pid = TryGetExcelProcessId(app);
            if (pid <= 0) return;

            if (IsAddinStartupCompleted(pid))
                return;

            WaitForVstoStartupByPid(pid, timeoutMs);
        }

        /// <summary>
        /// Excel.Application のメイン HWND からプロセス ID を取得する（終了確認・強制終了用）。
        /// </summary>
        public static int TryGetExcelProcessId(ExcelApp app)
        {
            if (app == null) return -1;

            IntPtr hwnd = IntPtr.Zero;
            try { hwnd = new IntPtr(app.Hwnd); } catch { return -1; }
            if (hwnd == IntPtr.Zero) return -1;

            try
            {
                uint pid;
                GetWindowThreadProcessId(hwnd, out pid);
                return (int)pid;
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// Quit 後も同一 PID が残る場合に VSTO が再ロードされないため、待機してから必要なら Kill する。
        /// </summary>
        public static void EnsureExcelProcessExited(int pid, int waitAfterQuitMs = 10000, int waitAfterKillMs = 5000, string logContext = null)
        {
            if (pid <= 0) return;

            if (WaitForProcessExitById(pid, waitAfterQuitMs))
                return;

            string ctx = string.IsNullOrEmpty(logContext) ? nameof(EnsureExcelProcessExited) : logContext;
            Debug.WriteLine($"[{ctx}] Excel PID {pid} still running after Quit; forcing kill.");
            TryKillProcessById(pid);
            WaitForProcessExitById(pid, waitAfterKillMs);
        }

        /// <summary>
        /// 名前が EXCEL のプロセスがいなくなるまで待つ（PID が取れない場合のフォールバック）。
        /// </summary>
        public static void WaitForAllExcelProcessesGone(int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                Process[] procs = null;
                try
                {
                    procs = Process.GetProcessesByName("EXCEL");
                    if (procs.Length == 0)
                        return;
                }
                finally
                {
                    if (procs != null)
                    {
                        foreach (var p in procs)
                        {
                            try { p.Dispose(); } catch { /* ignore */ }
                        }
                    }
                }

                Thread.Sleep(200);
            }
        }

        private static bool WaitForProcessExitById(int pid, int timeoutMs)
        {
            if (pid <= 0) return true;

            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    using (var p = Process.GetProcessById(pid))
                    {
                        if (p.HasExited)
                            return true;
                    }
                }
                catch (ArgumentException)
                {
                    return true;
                }

                Thread.Sleep(200);
            }

            try
            {
                using (var p = Process.GetProcessById(pid))
                {
                    return p.HasExited;
                }
            }
            catch (ArgumentException)
            {
                return true;
            }
        }

        private static void TryKillProcessById(int pid)
        {
            if (pid <= 0) return;

            try
            {
                using (var p = Process.GetProcessById(pid))
                {
                    if (!p.HasExited)
                        p.Kill();
                }
            }
            catch (ArgumentException)
            {
                /* already exited */
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExcelApplicationManager.TryKillProcessById] PID {pid}: {ex.Message}");
            }
        }

        private static void WaitForVstoStartupByPid(int excelPid, int timeoutMs)
        {
            if (excelPid <= 0) return;
            var sw = Stopwatch.StartNew();

            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (IsAddinStartupCompleted(excelPid))
                    return;

                Thread.Sleep(200);
            }
        }

        private static bool IsAddinStartupCompleted(int excelPid)
        {
            if (excelPid <= 0) return false;
            if (!File.Exists(AddinDiagPath)) return false;

            // 例: "[yyyy-MM-dd HH:mm:ss.xxx] [PID:12345] Startup completed"
            string pidToken = "[PID:" + excelPid.ToString() + "]";
            try
            {
                foreach (var line in File.ReadLines(AddinDiagPath))
                {
                    if (line == null) continue;
                    if (line.IndexOf(pidToken, StringComparison.OrdinalIgnoreCase) >= 0 &&
                        line.IndexOf(AddinStartupCompletedMarker, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        private static int StartExcelProcess()
        {
            string[] candidates = new[]
            {
                "excel.exe",
                @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
            };

            foreach (var path in candidates)
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = path,
                        // /e を付けると環境によって ROT 登録(=GetActiveObject)が遅延するため、
                        // 通常起動にして Application 取得を安定させる。
                        Arguments = "",
                        UseShellExecute = true
                    };
                    var proc = Process.Start(psi);
                    if (proc != null)
                    {
                        // ForegroundLock 制限の影響を下げるため、起動直後に Excel PID へ前面化許可を与える。
                        TryAllowExcelSetForeground(proc.Id);
                        return proc.Id;
                    }
                }
                catch
                {
                    // try next
                }
            }

            return -1;
        }

        private static void TryAllowExcelSetForeground(int pid)
        {
            if (pid <= 0) return;

            try
            {
                AllowSetForegroundWindow((uint)pid);
            }
            catch
            {
                // ignore
            }
        }
    }
}

