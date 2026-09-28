using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using Libraries;

namespace Ui.ViewModels
{
    /// <summary>
    /// Excel はすぐ見せる。アドインが記録を始めるまで、Excel のウィンドウを無効にし前面に案内を出す。
    /// </summary>
    internal static class ExcelStartupInputGate
    {
        const int TimeoutMs = 15000;
        const int PollMs = 200;

        static readonly Regex PidRegex = new Regex(@"\[PID:(\d+)\].*Startup completed", RegexOptions.Compiled);

        static readonly HashSet<int> ReadyPids = new HashSet<int>();
        static readonly List<IntPtr> DisabledWindows = new List<IntPtr>();
        static DispatcherTimer _timer;
        static Window _dialog;
        static Stopwatch _waiting;
        static bool _active;

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static void Begin()
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null)
                return;
            if (!dispatcher.CheckAccess())
            {
                dispatcher.BeginInvoke(new Action(Begin));
                return;
            }

            End();
            RefreshReadyPids();
            if (AllRunningExcelAddinsReady())
                return;

            _active = true;
            _waiting = Stopwatch.StartNew();
            _dialog = CreateDialog();
            _dialog.Show();
            DisableExcelWindows();

            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(PollMs) };
            _timer.Tick += OnTick;
            _timer.Start();
        }

        public static void End()
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
            {
                dispatcher.Invoke(new Action(End));
                return;
            }

            _active = false;
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

            foreach (IntPtr hwnd in DisabledWindows)
            {
                try { EnableWindow(hwnd, true); } catch { }
            }
            DisabledWindows.Clear();
        }

        static void OnTick(object sender, EventArgs e)
        {
            if (!_active)
                return;

            DisableExcelWindows();
            RefreshReadyPids();
            bool timedOut = _waiting != null && _waiting.ElapsedMilliseconds >= TimeoutMs;
            if (timedOut || AllRunningExcelAddinsReady())
                End();
        }

        static bool AllRunningExcelAddinsReady()
        {
            bool any = false;
            foreach (Process process in Process.GetProcessesByName("EXCEL"))
            {
                try
                {
                    if (process.HasExited)
                        continue;
                    any = true;
                    if (!ReadyPids.Contains(process.Id))
                        return false;
                }
                finally
                {
                    process.Dispose();
                }
            }
            return any;
        }

        static void RefreshReadyPids()
        {
            try
            {
                string path = ExcelLogReader.GetDiagnosticLogPath();
                if (!File.Exists(path))
                    return;
                foreach (string line in File.ReadLines(path))
                {
                    Match match = PidRegex.Match(line ?? "");
                    if (match.Success && int.TryParse(match.Groups[1].Value, out int pid))
                        ReadyPids.Add(pid);
                }
            }
            catch
            {
                /* 読めない間は待ちを続ける */
            }
        }

        static void DisableExcelWindows()
        {
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd))
                    return true;
                GetWindowThreadProcessId(hWnd, out uint pid);
                if (pid == 0)
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
