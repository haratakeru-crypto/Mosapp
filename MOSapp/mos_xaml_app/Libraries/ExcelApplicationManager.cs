using System;
using System.Diagnostics;
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

        public static ExcelApp GetOrCreateExcelApplication(bool makeVisible, int timeoutMs = 30000)
        {
            ExcelApp app = null;
            try
            {
                // 既に起動中なら生存確認済みのものを使う（終了直後の古い ROT は捨てる）
                app = TryGetHealthyExcelApplication();
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
        public static ExcelApp TryAttachRunningExcelApplication(bool makeVisible, int timeoutMs = 15000, bool waitForVstoStartup = true)
        {
            var app = WaitForActiveExcelApplication(timeoutMs);
            if (app == null)
                return null;

            TrySetVisible(app, makeVisible);
            if (!waitForVstoStartup)
                return app;

            int vstoWaitMs = Math.Min(timeoutMs, 8000);
            WaitForVstoStartupIfPossible(app, vstoWaitMs);
            return app;
        }

        /// <summary>
        /// ROT 上の Excel.Application を取得し、Hwnd まで応答する生存インスタンスだけ返す。
        /// 終了直後の古い登録（InvalidCastException / RPC 0x800706BA 等）は解放して null を返す。新規起動はしない。
        /// </summary>
        public static ExcelApp TryGetHealthyExcelApplication()
        {
            object raw = null;
            bool keep = false;
            try
            {
                raw = Marshal.GetActiveObject("Excel.Application");
                var app = raw as ExcelApp;
                if (app == null)
                    return null;

                // 古いプロキシはここで RPC 切断になる。直接キャストの例外を呼び出し元へ出さない。
                _ = app.Hwnd;
                keep = true;
                return app;
            }
            catch (COMException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (!keep)
                    ReleaseComObjectSafe(raw);
            }
        }

        private static ExcelApp WaitForActiveExcelApplication(int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                var app = TryGetHealthyExcelApplication();
                if (app != null) return app;
                Thread.Sleep(200);
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
                var app = TryGetHealthyExcelApplication();
                if (app != null)
                {
                    int pid = TryGetExcelProcessId(app);
                    if (pid == expectedPid)
                        return app;

                    // 別インスタンスや HWND 未確定の参照は保持せず、次の ROT 登録を待つ。
                    ReleaseComObjectSafe(app);
                }
                Thread.Sleep(200);
            }

            return null;
        }

        private static void ReleaseComObjectSafe(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
                return;

            try { Marshal.ReleaseComObject(comObject); } catch { }
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
        /// 起動中 Excel に VSTO Startup が無いとき、試験オープン前に終了させる（Word と同様）。
        /// LoadBehavior を直しても、既に動いている Excel にはアドインが載らないため。
        /// </summary>
        public static void RestartExcelIfAddInNotLoaded(string logContext = null)
        {
            ExcelApp app = null;
            int pid = -1;
            try
            {
                app = TryGetHealthyExcelApplication();
                if (app == null)
                    return;

                pid = TryGetExcelProcessId(app);
                if (pid > 0 && ExcelVstoReadiness.IsStartupCompleted(pid))
                    return;

                string ctx = string.IsNullOrEmpty(logContext) ? nameof(RestartExcelIfAddInNotLoaded) : logContext;
                ExcelVstoReadiness.RecordHostEvent(
                    "excel restart: no VSTO startup pid=" + pid + " context=" + ctx);
                try { app.DisplayAlerts = false; } catch { /* ignore */ }
                try { app.Quit(); } catch { /* ignore */ }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[RestartExcelIfAddInNotLoaded] {ex.Message}");
            }
            finally
            {
                ReleaseComObjectSafe(app);
            }

            if (pid > 0)
            {
                EnsureExcelProcessExited(
                    pid,
                    waitAfterQuitMs: 8000,
                    waitAfterKillMs: 4000,
                    logContext: logContext ?? nameof(RestartExcelIfAddInNotLoaded));
            }
            else
            {
                WaitForAllExcelProcessesGone(8000);
            }
        }

        /// <summary>
        /// Quit 後も同一 PID が残る場合に VSTO が再ロードされないため、待機してから必要なら Kill する。
        /// </summary>
        public static bool WaitForExcelProcessExit(int pid, int timeoutMs)
        {
            return WaitForProcessExitById(pid, timeoutMs);
        }

        public static void EnsureExcelProcessExited(int pid, int waitAfterQuitMs = 10000, int waitAfterKillMs = 5000, string logContext = null)
        {
            if (pid <= 0)
            {
                ExcelVstoReadiness.RecordHostEvent(
                    "excel-exit skip pid<=0 context=" + (logContext ?? nameof(EnsureExcelProcessExited)));
                return;
            }

            string ctx = string.IsNullOrEmpty(logContext) ? nameof(EnsureExcelProcessExited) : logContext;
            var sw = Stopwatch.StartNew();
            ExcelVstoReadiness.RecordHostEvent(
                "excel-exit begin context=" + ctx
                + " pid=" + pid
                + " waitQuitMs=" + waitAfterQuitMs
                + " waitKillMs=" + waitAfterKillMs
                + " alive=" + (IsProcessAlive(pid) ? "1" : "0")
                + " excelCount=" + CountExcelProcesses());

            if (WaitForProcessExitById(pid, waitAfterQuitMs))
            {
                ExcelVstoReadiness.RecordHostEvent(
                    "excel-exit ok context=" + ctx
                    + " pid=" + pid
                    + " how=wait-after-quit"
                    + " elapsed=" + sw.ElapsedMilliseconds + "ms"
                    + " excelCount=" + CountExcelProcesses());
                return;
            }

            Debug.WriteLine($"[{ctx}] Excel PID {pid} still running after Quit; forcing kill.");
            ExcelVstoReadiness.RecordHostEvent(
                "excel-exit kill context=" + ctx
                + " pid=" + pid
                + " afterQuitWaitMs=" + waitAfterQuitMs
                + " alive=" + (IsProcessAlive(pid) ? "1" : "0"));

            bool killOk = TryKillProcessById(pid);
            bool exitedAfterKill = WaitForProcessExitById(pid, waitAfterKillMs);
            ExcelVstoReadiness.RecordHostEvent(
                "excel-exit " + (exitedAfterKill ? "ok" : "FAILED")
                + " context=" + ctx
                + " pid=" + pid
                + " how=kill"
                + " killOk=" + (killOk ? "1" : "0")
                + " alive=" + (IsProcessAlive(pid) ? "1" : "0")
                + " elapsed=" + sw.ElapsedMilliseconds + "ms"
                + " excelCount=" + CountExcelProcesses());
        }

        /// <summary>
        /// 名前が EXCEL のプロセスがいなくなるまで待つ（PID が取れない場合のフォールバック）。
        /// </summary>
        public static void WaitForAllExcelProcessesGone(int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            ExcelVstoReadiness.RecordHostEvent(
                "excel-exit wait-all begin timeoutMs=" + timeoutMs
                + " excelCount=" + CountExcelProcesses());
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                Process[] procs = null;
                try
                {
                    procs = Process.GetProcessesByName("EXCEL");
                    if (procs.Length == 0)
                    {
                        ExcelVstoReadiness.RecordHostEvent(
                            "excel-exit wait-all ok elapsed=" + sw.ElapsedMilliseconds + "ms");
                        return;
                    }
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

            ExcelVstoReadiness.RecordHostEvent(
                "excel-exit wait-all FAILED timeoutMs=" + timeoutMs
                + " excelCount=" + CountExcelProcesses());
        }

        public static int CountExcelProcesses()
        {
            Process[] procs = null;
            try
            {
                procs = Process.GetProcessesByName("EXCEL");
                return procs.Length;
            }
            catch
            {
                return -1;
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
        }

        public static bool IsProcessAlive(int pid)
        {
            if (pid <= 0)
                return false;
            try
            {
                using (var p = Process.GetProcessById(pid))
                {
                    return !p.HasExited;
                }
            }
            catch (ArgumentException)
            {
                return false;
            }
            catch
            {
                return false;
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

        private static bool TryKillProcessById(int pid)
        {
            if (pid <= 0) return true;

            try
            {
                using (var p = Process.GetProcessById(pid))
                {
                    if (p.HasExited)
                        return true;
                    p.Kill();
                    return true;
                }
            }
            catch (ArgumentException)
            {
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExcelApplicationManager.TryKillProcessById] PID {pid}: {ex.Message}");
                ExcelVstoReadiness.RecordHostEvent(
                    "excel-exit kill-exception pid=" + pid + " error=" + ex.GetType().Name + ":" + ex.Message);
                return false;
            }
        }

        private static void WaitForVstoStartupByPid(int excelPid, int timeoutMs)
        {
            ExcelVstoReadiness.WaitForStartup(excelPid, timeoutMs);
        }

        private static bool IsAddinStartupCompleted(int excelPid)
        {
            return ExcelVstoReadiness.IsStartupCompleted(excelPid);
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

