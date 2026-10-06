using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using Libraries;

namespace Ui.ViewModels
{
    /// <summary>
    /// Excel はすぐ見せる。対象ブックのアドインが初期基準を終えるまで、その Excel だけを無効にし前面に案内を出す。
    /// </summary>
    internal static class ExcelStartupInputGate
    {
        const int TimeoutMs = 15000;
        const int PollMs = 200;

        static readonly List<IntPtr> DisabledWindows = new List<IntPtr>();
        static DispatcherTimer _timer;
        static Window _dialog;
        static Stopwatch _waiting;
        static bool _active;
        static int _targetPid;
        static string _expectedFile;
        static string _openToken;

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        [DllImport("user32.dll")]
        static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static void Begin(string expectedFilePath, string openToken)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null)
                return;
            if (!dispatcher.CheckAccess())
            {
                dispatcher.BeginInvoke(new Action(() => Begin(expectedFilePath, openToken)));
                return;
            }

            Finish(_active ? "restart" : null);
            _expectedFile = expectedFilePath ?? "";
            _openToken = openToken ?? "";
            _targetPid = 0;
            _active = true;
            _waiting = Stopwatch.StartNew();
            ExcelVstoReadiness.RecordHostEvent("gate start file=" + _expectedFile);
            _dialog = CreateDialog();
            _dialog.Show();
            DisableExcelWindows();

            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(PollMs) };
            _timer.Tick += OnTick;
            _timer.Start();
        }

        public static void NotifyTargetProcess(int processId)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null)
                return;
            if (!dispatcher.CheckAccess())
            {
                dispatcher.BeginInvoke(new Action(() => NotifyTargetProcess(processId)));
                return;
            }

            if (!_active || processId <= 0 || _targetPid == processId)
            {
                if (_active && processId > 0 && _targetPid == processId && IsTargetReady(processId))
                    Finish("ready");
                return;
            }

            _targetPid = processId;
            ExcelVstoReadiness.RecordHostEvent(
                "gate target-pid pid=" + processId
                + " elapsed=" + ElapsedMs()
                + "ms file=" + _expectedFile);
            // PID 確定前に無効化した他 Excel も含め、いったん全部戻してから対象だけ落とす。
            RestoreExcelInput(releaseTrackedOnly: false, onlyPid: 0);
            DisableExcelWindows();
            if (IsTargetReady(processId))
                Finish("ready");
        }

        public static void End()
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
            {
                dispatcher.Invoke(new Action(End));
                return;
            }

            Finish(_active ? "closed" : null);
        }

        /// <summary>
        /// ゲート外からも呼べる保険。無効のまま残った Excel トップレベルを有効に戻す。
        /// </summary>
        public static void EnsureExcelInputEnabled(int processId = 0)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
            {
                dispatcher.BeginInvoke(new Action(() => EnsureExcelInputEnabled(processId)));
                return;
            }

            RestoreExcelInput(releaseTrackedOnly: false, onlyPid: processId);
        }

        static void OnTick(object sender, EventArgs e)
        {
            if (!_active)
                return;

            DisableExcelWindows();
            bool ready = _targetPid > 0 && IsTargetReady(_targetPid);
            bool timedOut = _waiting != null && _waiting.ElapsedMilliseconds >= TimeoutMs;
            if (ready)
                Finish("ready");
            else if (timedOut)
                Finish("timeout");
        }

        static void Finish(string reason)
        {
            bool wasActive = _active;
            long elapsed = ElapsedMs();
            int pid = _targetPid;
            string file = _expectedFile;

            _active = false;
            _targetPid = 0;
            if (_timer != null)
            {
                _timer.Stop();
                _timer.Tick -= OnTick;
                _timer = null;
            }

            if (_dialog != null)
            {
                try { _dialog.Close(); } catch { }
                _dialog = null;
            }

            // 追跡漏れがあっても Excel を操作不能のまま残さない。
            RestoreExcelInput(releaseTrackedOnly: false, onlyPid: 0);
            _waiting = null;
            _expectedFile = null;
            _openToken = null;

            if (wasActive && !string.IsNullOrEmpty(reason))
            {
                ExcelVstoReadiness.RecordHostEvent(
                    "gate end reason=" + reason
                    + " elapsed=" + elapsed
                    + "ms pid=" + pid
                    + " file=" + file);
            }
        }

        static bool IsTargetReady(int processId)
        {
            return ExcelVstoReadiness.IsOpenReady(processId, _openToken);
        }

        static long ElapsedMs()
        {
            return _waiting == null ? 0 : _waiting.ElapsedMilliseconds;
        }

        static void RestoreExcelInput(bool releaseTrackedOnly, int onlyPid)
        {
            foreach (IntPtr hwnd in DisabledWindows)
            {
                try
                {
                    if (IsWindow(hwnd))
                        EnableWindow(hwnd, true);
                }
                catch { }
            }
            DisabledWindows.Clear();

            if (releaseTrackedOnly)
                return;

            ForceEnableExcelWindows(onlyPid);
        }

        static void ForceEnableExcelWindows(int onlyPid)
        {
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsExcelTopLevelWindow(hWnd, onlyPid))
                    return true;
                try
                {
                    if (!IsWindowEnabled(hWnd))
                        EnableWindow(hWnd, true);
                }
                catch { }
                return true;
            }, IntPtr.Zero);
        }

        static void DisableExcelWindows()
        {
            int onlyPid = _targetPid;
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsExcelTopLevelWindow(hWnd, onlyPid))
                    return true;

                if (DisabledWindows.Contains(hWnd))
                    return true;

                // EnableWindow の戻り値は「直前が有効だったか」。既に無効でも追跡する。
                try { EnableWindow(hWnd, false); } catch { }
                DisabledWindows.Add(hWnd);
                return true;
            }, IntPtr.Zero);
        }

        static bool IsExcelTopLevelWindow(IntPtr hWnd, int onlyPid)
        {
            if (!IsWindowVisible(hWnd))
                return false;
            GetWindowThreadProcessId(hWnd, out uint pid);
            if (pid == 0)
                return false;
            if (onlyPid > 0 && (int)pid != onlyPid)
                return false;
            try
            {
                using (Process process = Process.GetProcessById((int)pid))
                {
                    return string.Equals(process.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch
            {
                return false;
            }
        }

        static Window CreateDialog()
        {
            var dialog = new Window
            {
                Title = "準備中",
                Width = 420,
                Height = 160,
                WindowStyle = WindowStyle.None,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ShowInTaskbar = false,
                ResizeMode = ResizeMode.NoResize,
                Topmost = true,
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(30, 64, 175)),
                BorderThickness = new Thickness(2)
            };
            var stack = new StackPanel
            {
                Margin = new Thickness(24),
                VerticalAlignment = VerticalAlignment.Center
            };
            stack.Children.Add(new TextBlock
            {
                Text = "準備中です",
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(30, 64, 175)),
                HorizontalAlignment = HorizontalAlignment.Center
            });
            stack.Children.Add(new TextBlock
            {
                Text = "操作しないでください。準備が終わるまでお待ちください。",
                FontSize = 14,
                TextAlignment = TextAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 12, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center
            });
            dialog.Content = stack;
            return dialog;
        }
    }
}
