using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MOS_Word_app;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WordDoc = Microsoft.Office.Interop.Word.Document;
using WordProtectedViewWindow = Microsoft.Office.Interop.Word.ProtectedViewWindow;

namespace Libraries
{
    /// <summary>
    /// Word 取得と試験用ドキュメントオープン。
    /// 心拍が生きていれば即利用し、未準備時のみ再起動する（Excel に近い体感 + VSTO 保証）。
    /// </summary>
    public static class WordApplicationManager
    {
        private const int DefaultAcquireTimeoutMs = 10000;
        /// <summary>既存 Word で VSTO 起動直後を拾う短い待ち（従来 1500ms を短縮）。</summary>
        private const int ExistingWordHeartbeatWaitMs = 400;
        private const int VstoHeartbeatWaitAfterLaunchMs = 8000;
        private const int WordQuitWaitMs = 5000;
        private const int HeartbeatPollIntervalMs = 100;
        /// <summary>VSTO が ShowAll ポーリング中なら数秒以内に更新される。</summary>
        private const int ActiveVstoHeartbeatMaxAgeSeconds = 15;
        private const int OpenRetryCount = 8;
        private const int OpenRetryDelayMs = 250;
        private const int RpcCallRejected = unchecked((int)0x80010001);
        private const int RpcServerCallRetryLater = unchecked((int)0x8001010A);

        private const int SwHide = 0;
        private const int SwShow = 5;
        private const int SwRestore = 9;

        [DllImport("user32.dll")]
        private static extern bool AllowSetForegroundWindow(uint dwProcessId);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        /// <summary>
        /// 試験用: Release VSTO を有効化し、ハートビートが取れる Word を返す（既存 Word に VSTO が無ければ再起動）。
        /// </summary>
        /// <param name="preferredDocumentPath">
        /// コールド起動時に指定すると、空起動(/n)ではなく文書付きで起動し、VSTO 読込と Open を並行化できる。
        /// </param>
        public static WordApp AcquireWordApplicationForExam(
            bool makeVisible = true,
            int timeoutMs = DefaultAcquireTimeoutMs,
            string preferredDocumentPath = null)
        {
            var sw = Stopwatch.StartNew();
            VSTOInstallerHelper.EnsureAddInReadyForExam(out _);

            WordApp existing = TryGetActiveWordApplication();
            if (existing != null)
            {
                TrySetVisible(existing, makeVisible);

                // 既に VSTO が動いていれば待ちなし
                if (LogReader.IsVstoHeartbeatFresh(ActiveVstoHeartbeatMaxAgeSeconds))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordApplicationManager] Acquire fast-path (active VSTO) {sw.ElapsedMilliseconds}ms");
                    return existing;
                }

                // 起動直後の短い窓だけ待つ（従来の 1.5s 固定待ちはしない）
                if (WaitForVstoHeartbeat(ExistingWordHeartbeatWaitMs))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordApplicationManager] Acquire existing after brief wait {sw.ElapsedMilliseconds}ms");
                    return existing;
                }

                System.Diagnostics.Debug.WriteLine(
                    "[WordApplicationManager] Existing Word has no VSTO heartbeat; restarting.");
                TryQuitWordAndWait(existing);
                try { Marshal.FinalReleaseComObject(existing); } catch { /* ignore */ }
            }

            LogReader.ClearVstoHeartbeat();
            // 文書付き起動を維持（VSTO と Open を並行）。makeVisible=false なら掴み次第非表示にする。
            WordApp launched = LaunchWordAndWaitForVsto(makeVisible, timeoutMs, preferredDocumentPath);
            System.Diagnostics.Debug.WriteLine(
                $"[WordApplicationManager] Acquire cold-start {sw.ElapsedMilliseconds}ms visible={makeVisible}");
            return launched;
        }

        /// <summary>
        /// Word を用意して試験ドキュメントを開く。失敗時は false（試験バーへ進まない用途）。
        /// オープン前に Zone.Identifier を外し、保護ビューなら Edit() で解除する。
        /// </summary>
        public static bool TryOpenExamDocument(string filePath, out WordApp app, bool makeVisible = true)
        {
            app = null;
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return false;

            var sw = Stopwatch.StartNew();
            try
            {
                // 保護ビュー予防（インターネット由来マーク）
                WordDataPathHelper.RemoveZoneIdentifier(filePath);

                app = AcquireWordApplicationForExam(makeVisible, DefaultAcquireTimeoutMs, filePath);
                if (app == null)
                    return false;

                if (TryEnsureDocumentEditableAndActive(app, filePath))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordApplicationManager] TryOpenExamDocument ready {sw.ElapsedMilliseconds}ms");
                    return true;
                }

                bool opened = TryOpenDocumentWithRetry(app, filePath);
                System.Diagnostics.Debug.WriteLine(
                    $"[WordApplicationManager] TryOpenExamDocument open={(opened ? "ok" : "fail")} {sw.ElapsedMilliseconds}ms file={Path.GetFileName(filePath)}");
                return opened;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WordApplicationManager] TryOpenExamDocument error {sw.ElapsedMilliseconds}ms: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 起動済み Word のまま次の試験文書を開く。Word の終了・再起動・起動待ちはしない。
        /// 既存 Word が無い、または VSTO 心拍が古いときは false。
        /// </summary>
        public static Task<bool> TrySwitchToDocumentInRunningWordAsync(string filePath)
        {
            return RunOnStaAsync(() => TrySwitchToDocumentInRunningWord(filePath));
        }

        /// <summary>
        /// 初回起動相当のオープンを STA スレッドで行う。次プロジェクトの失敗時だけ使う。
        /// </summary>
        public static Task<bool> TryOpenExamDocumentOnStaAsync(string filePath, bool makeVisible = true)
        {
            return RunOnStaAsync(() =>
            {
                bool opened = TryOpenExamDocument(filePath, out WordApp app, makeVisible);
                if (app != null)
                {
                    try { Marshal.ReleaseComObject(app); } catch { /* ignore */ }
                }
                return opened;
            });
        }

        private static Task<bool> RunOnStaAsync(Func<bool> action)
        {
            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var thread = new Thread(() =>
            {
                try
                {
                    tcs.TrySetResult(action());
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[WordApplicationManager] STA worker: " + ex.Message);
                    tcs.TrySetResult(false);
                }
            });
            thread.IsBackground = true;
            thread.Name = "WordExamSta";
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return tcs.Task;
        }

        private static bool TrySwitchToDocumentInRunningWord(string filePath)
        {
            var sw = Stopwatch.StartNew();
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                System.Diagnostics.Debug.WriteLine("[WordApplicationManager] Next-project fast-path skipped: file missing");
                return false;
            }

            if (!LogReader.IsVstoHeartbeatFresh(ActiveVstoHeartbeatMaxAgeSeconds))
            {
                System.Diagnostics.Debug.WriteLine("[WordApplicationManager] Next-project fast-path skipped: heartbeat stale");
                return false;
            }

            WordApp wordApp = TryGetActiveWordApplication();
            if (wordApp == null)
            {
                System.Diagnostics.Debug.WriteLine("[WordApplicationManager] Next-project fast-path skipped: no Word");
                return false;
            }

            try
            {
                WordDataPathHelper.RemoveZoneIdentifier(filePath);
                try
                {
                    wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                }
                catch { /* ignore */ }

                TrySetVisible(wordApp, true);
                LogReader.RequestCloseNavigationPaneIfOpen();
                if (!TrySaveAndCloseOpenDocuments(wordApp))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordApplicationManager] Next-project fast-path close failed {sw.ElapsedMilliseconds}ms");
                    return false;
                }

                bool opened = TryOpenDocumentWithRetry(wordApp, filePath);
                System.Diagnostics.Debug.WriteLine(
                    $"[WordApplicationManager] Next-project fast-path open={(opened ? "ok" : "fail")} {sw.ElapsedMilliseconds}ms file={Path.GetFileName(filePath)}");
                return opened;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WordApplicationManager] Next-project fast-path error {sw.ElapsedMilliseconds}ms: {ex.Message}");
                return false;
            }
            finally
            {
                try { Marshal.ReleaseComObject(wordApp); } catch { /* ignore */ }
            }
        }

        private static bool TrySaveAndCloseOpenDocuments(WordApp wordApp)
        {
            if (wordApp == null)
                return false;

            try
            {
                int guard = 0;
                try { guard = wordApp.Documents.Count + 2; } catch { return false; }

                while (guard-- > 0)
                {
                    int count;
                    try { count = wordApp.Documents.Count; }
                    catch { return false; }
                    if (count <= 0)
                        return true;

                    WordDoc doc = null;
                    try
                    {
                        doc = wordApp.Documents[1];
                        try
                        {
                            if (doc.Saved == false)
                                doc.Save();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                "[WordApplicationManager] Next-project save: " + ex.Message);
                        }

                        doc.Close(SaveChanges: false);
                    }
                    catch (COMException ex) when (ex.HResult == unchecked((int)0x80010108))
                    {
                        System.Diagnostics.Debug.WriteLine("[WordApplicationManager] Next-project close disconnected");
                        return false;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            "[WordApplicationManager] Next-project close: " + ex.Message);
                        return false;
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try { Marshal.ReleaseComObject(doc); } catch { /* ignore */ }
                        }
                    }
                }

                try
                {
                    return wordApp.Documents.Count == 0;
                }
                catch
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[WordApplicationManager] Next-project close: " + ex.Message);
                return false;
            }
        }

        /// <summary>同じパスの文書が開いていれば保存して閉じる（保護ビュー含む）。</summary>
        public static void TryCloseOpenDocumentByPath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return;

            WordApp wordApp = TryGetActiveWordApplication();
            if (wordApp == null)
                return;

            string target = NormalizePath(filePath);
            TryCloseProtectedViewByPath(wordApp, target);

            try
            {
                for (int i = wordApp.Documents.Count; i >= 1; i--)
                {
                    WordDoc doc = wordApp.Documents[i];
                    try
                    {
                        if (!PathsEqual(doc.FullName, target))
                            continue;

                        try
                        {
                            if (!doc.Saved)
                                doc.Save();
                        }
                        catch { /* ignore save failure; still try close */ }

                        doc.Close(SaveChanges: false);
                        break;
                    }
                    finally
                    {
                        try { if (doc != null) Marshal.ReleaseComObject(doc); } catch { /* ignore */ }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordApplicationManager] TryCloseOpenDocumentByPath: {ex.Message}");
            }
        }

        public static bool WaitForVstoHeartbeat(int timeoutMs)
        {
            if (timeoutMs <= 0)
                return LogReader.IsVstoHeartbeatFresh(ActiveVstoHeartbeatMaxAgeSeconds);

            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (LogReader.IsVstoHeartbeatFresh(ActiveVstoHeartbeatMaxAgeSeconds))
                    return true;
                Thread.Sleep(HeartbeatPollIntervalMs);
            }

            return LogReader.IsVstoHeartbeatFresh(ActiveVstoHeartbeatMaxAgeSeconds);
        }

        /// <summary>RPC 拒否などに対して Documents.Open をリトライする。</summary>
        public static bool TryOpenDocumentWithRetry(WordApp app, string filePath, int maxAttempts = OpenRetryCount)
        {
            if (app == null || string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return false;

            if (TryEnsureDocumentEditableAndActive(app, filePath))
                return true;

            Exception last = null;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    app.Documents.Open(filePath, ReadOnly: false, Visible: true);
                    return TryEnsureDocumentEditableAndActive(app, filePath);
                }
                catch (COMException ex) when (IsRetryableComError(ex) && attempt < maxAttempts)
                {
                    last = ex;
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordApplicationManager] Documents.Open retry {attempt}/{maxAttempts}: 0x{ex.HResult:X8}");
                    // 保護ビュー中は Open が拒否されやすいので毎回解除を試す
                    if (TryEnsureDocumentEditableAndActive(app, filePath))
                        return true;
                    Thread.Sleep(OpenRetryDelayMs * attempt);
                }
                catch (COMException ex)
                {
                    last = ex;
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordApplicationManager] Documents.Open COM failed: 0x{ex.HResult:X8} {ex.Message}");
                    if (TryEnsureDocumentEditableAndActive(app, filePath))
                        return true;
                    break;
                }
                catch (Exception ex)
                {
                    last = ex;
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordApplicationManager] Documents.Open failed: {ex.Message}");
                    if (TryEnsureDocumentEditableAndActive(app, filePath))
                        return true;
                    break;
                }
            }

            if (last != null)
                System.Diagnostics.Debug.WriteLine($"[WordApplicationManager] Open gave up: {last.Message}");
            return TryEnsureDocumentEditableAndActive(app, filePath);
        }

        /// <summary>
        /// 通常文書としてアクティブ化する。保護ビューなら Edit() で解除してからアクティブ化する。
        /// </summary>
        public static bool TryEnsureDocumentEditableAndActive(WordApp wordApp, string filePath)
        {
            if (TryActivateOpenDocument(wordApp, filePath))
                return true;

            if (!TryExitProtectedView(wordApp, filePath))
                return false;

            // Edit() 成功後、Documents 側に載るまで短い待ちを入れる
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < 1500)
            {
                if (TryActivateOpenDocument(wordApp, filePath))
                    return true;
                Thread.Sleep(100);
            }

            // Edit 自体は成功しているので、Activate できなくても開いた扱いにする
            return true;
        }

        public static bool TryActivateOpenDocument(WordApp wordApp, string filePath)
        {
            if (wordApp == null || string.IsNullOrWhiteSpace(filePath))
                return false;

            string target = NormalizePath(filePath);
            if (string.IsNullOrEmpty(target))
                return false;

            try
            {
                for (int i = wordApp.Documents.Count; i >= 1; i--)
                {
                    WordDoc doc = wordApp.Documents[i];
                    try
                    {
                        if (!PathsEqual(doc.FullName, target))
                            continue;

                        doc.Activate();
                        try { doc.ActiveWindow?.Activate(); } catch { /* ignore */ }
                        return true;
                    }
                    finally
                    {
                        try { if (doc != null) Marshal.ReleaseComObject(doc); } catch { /* ignore */ }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordApplicationManager] TryActivateOpenDocument: {ex.Message}");
            }

            return false;
        }

        /// <summary>保護ビューの対象ファイルを Edit() で通常編集モードへ切り替える。</summary>
        public static bool TryExitProtectedView(WordApp wordApp, string filePath)
        {
            if (wordApp == null || string.IsNullOrWhiteSpace(filePath))
                return false;

            string target = NormalizePath(filePath);
            string targetFileName = Path.GetFileName(filePath);

            try
            {
                // Count が変わるため末尾から
                for (int i = wordApp.ProtectedViewWindows.Count; i >= 1; i--)
                {
                    WordProtectedViewWindow pvw = null;
                    try
                    {
                        pvw = wordApp.ProtectedViewWindows[i];
                        if (!ProtectedViewMatches(pvw, target, targetFileName))
                            continue;

                        System.Diagnostics.Debug.WriteLine(
                            $"[WordApplicationManager] Protected view detected, calling Edit(): {Path.GetFileName(filePath)}");
                        WordDoc doc = pvw.Edit();
                        try
                        {
                            if (doc != null)
                            {
                                try { doc.Activate(); } catch { /* ignore */ }
                                try { doc.ActiveWindow?.Activate(); } catch { /* ignore */ }
                            }
                        }
                        finally
                        {
                            try { if (doc != null) Marshal.ReleaseComObject(doc); } catch { /* ignore */ }
                        }

                        return true;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[WordApplicationManager] ProtectedView.Edit failed: {ex.Message}");
                    }
                    finally
                    {
                        try { if (pvw != null) Marshal.ReleaseComObject(pvw); } catch { /* ignore */ }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordApplicationManager] TryExitProtectedView: {ex.Message}");
            }

            return false;
        }

        private static bool ProtectedViewMatches(WordProtectedViewWindow pvw, string targetFullPath, string targetFileName)
        {
            if (pvw == null)
                return false;

            try
            {
                string sourcePath = (pvw.SourcePath ?? "").Trim();
                string sourceName = (pvw.SourceName ?? "").Trim();
                if (!string.IsNullOrEmpty(sourcePath) && !string.IsNullOrEmpty(sourceName))
                {
                    string combined = NormalizePath(Path.Combine(sourcePath, sourceName));
                    if (PathsEqual(combined, targetFullPath))
                        return true;
                }

                if (!string.IsNullOrEmpty(sourceName)
                    && sourceName.Equals(targetFileName, StringComparison.OrdinalIgnoreCase))
                    return true;

                string caption = (pvw.Caption ?? "").Trim();
                if (!string.IsNullOrEmpty(caption)
                    && (caption.Equals(targetFileName, StringComparison.OrdinalIgnoreCase)
                        || caption.IndexOf(targetFileName, StringComparison.OrdinalIgnoreCase) >= 0))
                    return true;
            }
            catch
            {
                return false;
            }

            return false;
        }

        private static void TryCloseProtectedViewByPath(WordApp wordApp, string targetFullPath)
        {
            if (wordApp == null || string.IsNullOrEmpty(targetFullPath))
                return;

            string targetFileName = Path.GetFileName(targetFullPath);
            try
            {
                for (int i = wordApp.ProtectedViewWindows.Count; i >= 1; i--)
                {
                    WordProtectedViewWindow pvw = null;
                    try
                    {
                        pvw = wordApp.ProtectedViewWindows[i];
                        if (!ProtectedViewMatches(pvw, targetFullPath, targetFileName))
                            continue;
                        pvw.Close();
                        break;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[WordApplicationManager] Close protected view: {ex.Message}");
                    }
                    finally
                    {
                        try { if (pvw != null) Marshal.ReleaseComObject(pvw); } catch { /* ignore */ }
                    }
                }
            }
            catch
            {
                // ignore
            }
        }

        private static bool IsRetryableComError(COMException ex)
        {
            return ex.HResult == RpcCallRejected || ex.HResult == RpcServerCallRetryLater;
        }

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return string.Empty;
            try { return Path.GetFullPath(path); }
            catch { return path; }
        }

        private static bool PathsEqual(string left, string rightNormalized)
        {
            if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(rightNormalized))
                return false;
            return string.Equals(NormalizePath(left), rightNormalized, StringComparison.OrdinalIgnoreCase);
        }

        private static WordApp LaunchWordAndWaitForVsto(bool makeVisible, int timeoutMs, string preferredDocumentPath)
        {
            int launchedPid = StartWordProcess(preferredDocumentPath);
            if (launchedPid <= 0)
                throw new InvalidOperationException("Word を起動できませんでした。");

            WordApp app = WaitForActiveWordApplication(timeoutMs);
            if (app == null)
                throw new InvalidOperationException("起動後の Word へ接続できませんでした。");

            // 文書付き起動でも、掴めた直後に表示方針へ合わせる。非表示指定のときだけ隠す。
            TrySetVisible(app, makeVisible);
            if (!makeVisible)
                KeepWordHidden(app);
            else
                MOS_Word_app.Views.WordStartupInputGate.DisableWordWindows();

            // 文書付き起動時は VSTO と文書読込が並行するため、心拍は短めに待ちつつ Open リトライに委ねる
            int vstoWaitMs = string.IsNullOrEmpty(preferredDocumentPath)
                ? Math.Min(timeoutMs, VstoHeartbeatWaitAfterLaunchMs)
                : Math.Min(timeoutMs, 3000);

            if (!WaitForVstoHeartbeat(vstoWaitMs))
                System.Diagnostics.Debug.WriteLine("[WordApplicationManager] VSTO heartbeat not detected after Word launch.");

            // 文書付き起動時は Open / 保護ビュー解除完了を短く待つ
            if (!string.IsNullOrWhiteSpace(preferredDocumentPath))
            {
                var appear = Stopwatch.StartNew();
                while (appear.ElapsedMilliseconds < 2500
                    && !TryEnsureDocumentEditableAndActive(app, preferredDocumentPath))
                {
                    if (!makeVisible)
                        KeepWordHidden(app);
                    else
                        MOS_Word_app.Views.WordStartupInputGate.DisableWordWindows();
                    Thread.Sleep(100);
                }
            }

            if (!makeVisible)
                KeepWordHidden(app);

            return app;
        }

        private static void KeepWordHidden(WordApp app)
        {
            if (app == null) return;
            try { app.Visible = false; } catch { /* ignore */ }
            HideWordHwnds();
        }

        private static void HideWordHwnds()
        {
            try
            {
                foreach (Process p in Process.GetProcessesByName("WINWORD"))
                {
                    try
                    {
                        uint pid = (uint)p.Id;
                        EnumWindows((hWnd, lParam) =>
                        {
                            GetWindowThreadProcessId(hWnd, out uint windowPid);
                            if (windowPid != pid)
                                return true;
                            var sb = new StringBuilder(64);
                            GetClassName(hWnd, sb, sb.Capacity);
                            if (sb.ToString() == "OpusApp")
                                ShowWindow(hWnd, SwHide);
                            return true;
                        }, IntPtr.Zero);
                    }
                    catch { /* ignore */ }
                    finally
                    {
                        try { p.Dispose(); } catch { /* ignore */ }
                    }
                }
            }
            catch { /* ignore */ }
        }

        private static void ShowWordHwnds()
        {
            try
            {
                foreach (Process p in Process.GetProcessesByName("WINWORD"))
                {
                    try
                    {
                        uint pid = (uint)p.Id;
                        EnumWindows((hWnd, lParam) =>
                        {
                            GetWindowThreadProcessId(hWnd, out uint windowPid);
                            if (windowPid != pid)
                                return true;
                            var sb = new StringBuilder(64);
                            GetClassName(hWnd, sb, sb.Capacity);
                            if (sb.ToString() == "OpusApp")
                            {
                                ShowWindow(hWnd, SwRestore);
                                ShowWindow(hWnd, SwShow);
                            }
                            return true;
                        }, IntPtr.Zero);
                    }
                    catch { /* ignore */ }
                    finally
                    {
                        try { p.Dispose(); } catch { /* ignore */ }
                    }
                }
            }
            catch { /* ignore */ }
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
                    try { p.Dispose(); } catch { /* ignore */ }
                }
            }
        }

        /// <summary>
        /// 模試アプリ終了時: 未保存文書を保存して Word を終了し、残プロセスがあれば待機／強制終了する。
        /// サインアウト阻害のゾンビ COM / WINWORD 残りを防ぐ。
        /// </summary>
        public static void CloseWordApplicationForAppExit()
        {
            WordApp app = null;
            try
            {
                app = TryGetActiveWordApplication();
                if (app == null)
                    return;

                try
                {
                    app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                }
                catch { /* ignore */ }

                try
                {
                    for (int i = app.Documents.Count; i >= 1; i--)
                    {
                        WordDoc doc = null;
                        try
                        {
                            doc = app.Documents[i];
                            if (doc.Saved == false)
                                doc.Save();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                "[CloseWordApplicationForAppExit] Save: " + ex.Message);
                        }
                        finally
                        {
                            if (doc != null)
                            {
                                try { Marshal.ReleaseComObject(doc); } catch { /* ignore */ }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        "[CloseWordApplicationForAppExit] Documents: " + ex.Message);
                }

                TryQuitWordAndWait(app);
                app = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[CloseWordApplicationForAppExit] " + ex.Message);
            }
            finally
            {
                if (app != null)
                {
                    try { Marshal.ReleaseComObject(app); } catch { /* ignore */ }
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
            try { app.Visible = makeVisible; } catch { /* ignore */ }
            if (makeVisible)
                ShowWordHwnds();
            else
                HideWordHwnds();
        }

        /// <summary>準備中ダイアログを閉じたあとに Word を表示する。</summary>
        public static void SetWordVisible(bool visible)
        {
            TrySetVisible(TryGetActiveWordApplication(), visible);
        }

        private static int StartWordProcess(string documentPath = null)
        {
            string[] candidates =
            {
                "winword.exe",
                @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
            };

            string args;
            if (!string.IsNullOrWhiteSpace(documentPath) && File.Exists(documentPath))
                args = "\"" + documentPath.Replace("\"", "") + "\"";
            else
                args = "/n"; // 文書を開かずに起動（空の Document1 を作らない）

            foreach (string path in candidates)
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = path,
                        Arguments = args,
                        UseShellExecute = true
                    };
                    Process proc = Process.Start(psi);
                    if (proc == null)
                        continue;

                    try { AllowSetForegroundWindow((uint)proc.Id); } catch { /* ignore */ }
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
