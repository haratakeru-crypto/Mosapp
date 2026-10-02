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
        static extern bool IsWindowVisible(IntPtr hWnd);

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
            ReleaseDisabledWindows();
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

            ReleaseDisabledWindows();
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

        static void ReleaseDisabledWindows()
        {
            foreach (IntPtr hwnd in DisabledWindows)
            {
                try { EnableWindow(hwnd, true); } catch { }
            }
            DisabledWindows.Clear();
        }

        static void DisableExcelWindows()
        {
            int onlyPid = _targetPid;
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd))
                    return true;
                GetWindowThreadProcessId(hWnd, out uint pid);
                if (pid == 0)
                    return true;
                if (onlyPid > 0 && (int)pid != onlyPid)
                    return true;
                try
                {
                    using (Process process = Process.GetProcessById((int)pid))
                    {
                        if (!string.Equals(process.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
                catch
                {
                    return true;
                }

                if (DisabledWindows.Contains(hWnd))
                    return true;
                if (EnableWindow(hWnd, false))
                    DisabledWindows.Add(hWnd);
                return true;
            }, IntPtr.Zero);
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
