using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using Libraries;

namespace MOS_PowerPoint_app.Views
{
    /// <summary>
    /// PowerPoint はすぐ見せる。アドインが記録を始めるまで、PowerPoint のウィンドウを無効にし前面に案内を出す。
    /// </summary>
    public static class PowerPointStartupInputGate
    {
        const int TimeoutMs = 15000;
        const int PollMs = 200;
        const int HeartbeatMaxAgeSeconds = 15;

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
        static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static void Begin()
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null)
                return;
            if (!dispatcher.CheckAccess())
            {
                dispatcher.Invoke(new Action(Begin));
                return;
            }

            Finish(_active);
            // 残っている心拍ファイルだけでは足りない。PowerPoint が動いていて心拍が新しいときだけ案内を出さない。
            if (IsPowerPointRunning() && PPLogReader.IsVstoHeartbeatFresh(HeartbeatMaxAgeSeconds))
                return;

            _active = true;
            _waiting = Stopwatch.StartNew();
            _dialog = CreateDialog();
            _dialog.Show();
            DisablePowerPointWindows();

            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(PollMs) };
            _timer.Tick += OnTick;
            _timer.Start();
        }

        public static void End()
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
            {
                try { dispatcher.Invoke(new Action(End)); } catch { }
                return;
            }

            Finish(_active);
        }

        /// <summary>
        /// ゲート外からも呼べる保険。無効のまま残った PowerPoint トップレベルを有効に戻す。
        /// </summary>
        public static void EnsurePowerPointInputEnabled()
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
            {
                dispatcher.BeginInvoke(new Action(EnsurePowerPointInputEnabled));
                return;
            }

            RestorePowerPointInput(releaseTrackedOnly: false);
        }

        public static void DisablePowerPointWindows()
        {
            if (!_active)
                return;

            EnumWindows((hWnd, lParam) =>
            {
                if (!IsPowerPointTopLevelWindow(hWnd))
                    return true;

                if (DisabledWindows.Contains(hWnd))
                    return true;

                // EnableWindow の戻り値は「直前が有効だったか」。既に無効でも追跡する。
                try { EnableWindow(hWnd, false); } catch { }
                DisabledWindows.Add(hWnd);
                return true;
            }, IntPtr.Zero);
        }

        static void Finish(bool wasActive)
        {
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

            // 追跡漏れがあっても PowerPoint を操作不能のまま残さない。
            if (wasActive || DisabledWindows.Count > 0)
                RestorePowerPointInput(releaseTrackedOnly: false);

            _waiting = null;
        }

        static void OnTick(object sender, EventArgs e)
        {
            if (!_active)
                return;

            DisablePowerPointWindows();
            bool timedOut = _waiting != null && _waiting.ElapsedMilliseconds >= TimeoutMs;
            bool addInReady = IsPowerPointRunning() && PPLogReader.IsVstoHeartbeatFresh(HeartbeatMaxAgeSeconds);
            if (timedOut || addInReady)
                End();
        }

        static void RestorePowerPointInput(bool releaseTrackedOnly)
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

            ForceEnablePowerPointWindows();
        }

        static void ForceEnablePowerPointWindows()
        {
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsPowerPointTopLevelWindow(hWnd))
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

        static bool IsPowerPointTopLevelWindow(IntPtr hWnd)
        {
            if (!IsWindowVisible(hWnd))
                return false;
            GetWindowThreadProcessId(hWnd, out uint pid);
            if (pid == 0)
                return false;
            try
            {
                using (Process process = Process.GetProcessById((int)pid))
                {
                    return string.Equals(process.ProcessName, "POWERPNT", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch
            {
                return false;
            }
        }

        static bool IsPowerPointRunning()
        {
            Process[] processes = Process.GetProcessesByName("POWERPNT");
            try
            {
                foreach (Process process in processes)
                {
                    try
                    {
                        if (!process.HasExited)
                            return true;
                    }
                    catch { }
                }
                return false;
            }
            finally
            {
                foreach (Process process in processes)
                {
                    try { process.Dispose(); } catch { }
                }
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
