using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using MOSExcelMogiApp.Infrastructure;

namespace MOSExcelMogiApp
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        const string SingleInstanceMutexName = "Local\\MOS_Excel_app_SingleInstance";
        const string ActivateEventName = "Local\\MOS_Excel_app_Activate";
        const int ActivateSignalWaitMs = 500;

        static Mutex _singleInstanceMutex;
        static EventWaitHandle _activateEvent;
        static CancellationTokenSource _activateListenCts;
        static Thread _activateListenThread;
        Window _startupSplash;

        protected override void OnStartup(StartupEventArgs e)
        {
            bool createdNew;
            try
            {
                _singleInstanceMutex = new Mutex(true, SingleInstanceMutexName, out createdNew);
            }
            catch
            {
                createdNew = true;
                _singleInstanceMutex = null;
            }

            if (!createdNew)
            {
                try { _singleInstanceMutex?.Dispose(); } catch { }
                _singleInstanceMutex = null;

                bool signaled = SignalExistingInstanceToActivate();
                bool activated = TryActivateExistingInstance();
                if (!signaled && !activated)
                {
                    MessageBox.Show(
                        "MOS Excel アプリは既に起動しています。\n起動中の場合は、しばらくお待ちください。",
                        "MOS Excel",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }

                Shutdown();
                return;
            }

            StartActivateListener();

            _startupSplash = CreateStartupSplash();
            _startupSplash.Show();
            // 連打時に「起動中」が見えるよう、一度描画を進める
            DoEvents(_startupSplash.Dispatcher);

            base.OnStartup(e);

            var main = new MainWindow();
            MainWindow = main;
            main.Show();
            ShutdownMode = ShutdownMode.OnMainWindowClose;

            CloseStartupSplash();

            // MainWindow 表示後に VSTO 準備と類題 xlsx 先読み（UI スレッドをブロックしない）
            Dispatcher.BeginInvoke(new Action(() =>
            {
                Libraries.ExcelVstoInstallerHelper.StartBackgroundPrepForExam();
                Task.Run(() => VariantProblemWorkbook.Preload());
            }), DispatcherPriority.ApplicationIdle);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            CloseStartupSplash();
            StopActivateListener();
            ReleaseSingleInstanceMutex();
            base.OnExit(e);
        }

        private static void DoEvents(Dispatcher dispatcher)
        {
            if (dispatcher == null)
                return;
            try
            {
                dispatcher.Invoke(new Action(() => { }), DispatcherPriority.Render);
            }
            catch { /* ignore */ }
        }

        private static Window CreateStartupSplash()
        {
            var title = new TextBlock
            {
                Text = "起動中です",
                FontSize = 20,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(0x1E, 0x3A, 0x5F)),
                HorizontalAlignment = HorizontalAlignment.Center
            };
            var message = new TextBlock
            {
                Text = "MOS Excel アプリを起動しています。\nしばらくお待ちください。",
                FontSize = 14,
                Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
                TextAlignment = TextAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 12, 0, 0)
            };
            var panel = new StackPanel();
            panel.Children.Add(title);
            panel.Children.Add(message);

            return new Window
            {
                Title = "MOS Excel",
                Width = 420,
                Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.None,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = true,
                Topmost = true,
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0x1E, 0x3A, 0x5F)),
                BorderThickness = new Thickness(2),
                Content = new Border
                {
                    Padding = new Thickness(28, 28, 28, 28),
                    Child = panel
                }
            };
        }

        private void CloseStartupSplash()
        {
            try
            {
                if (_startupSplash != null)
                {
                    _startupSplash.Close();
                    _startupSplash = null;
                }
            }
            catch { /* ignore */ }
        }

        private static void ReleaseSingleInstanceMutex()
        {
            if (_singleInstanceMutex == null)
                return;
            try
            {
                _singleInstanceMutex.ReleaseMutex();
            }
            catch { /* ignore */ }
            try
            {
                _singleInstanceMutex.Dispose();
            }
            catch { /* ignore */ }
            _singleInstanceMutex = null;
        }

        /// <summary>既存インスタンスへ「UI を前面へ」を通知する。通知できれば true。</summary>
        private static bool SignalExistingInstanceToActivate()
        {
            EventWaitHandle ev = null;
            try
            {
                bool createdNew;
                ev = new EventWaitHandle(false, EventResetMode.AutoReset, ActivateEventName, out createdNew);
                if (createdNew)
                    return false;

                ev.Set();
                Thread.Sleep(ActivateSignalWaitMs);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[App] SignalExistingInstanceToActivate: " + ex.Message);
                return false;
            }
            finally
            {
                try { ev?.Dispose(); } catch { /* ignore */ }
            }
        }

        void StartActivateListener()
        {
            try
            {
                bool createdNew;
                _activateEvent = new EventWaitHandle(false, EventResetMode.AutoReset, ActivateEventName, out createdNew);
                _activateListenCts = new CancellationTokenSource();
                CancellationToken token = _activateListenCts.Token;
                _activateListenThread = new Thread(() =>
                {
                    WaitHandle[] handles = { _activateEvent, token.WaitHandle };
                    while (!token.IsCancellationRequested)
                    {
                        int index;
                        try
                        {
                            index = WaitHandle.WaitAny(handles);
                        }
                        catch
                        {
                            break;
                        }

                        if (token.IsCancellationRequested || index != 0)
                            break;

                        try
                        {
                            Dispatcher.BeginInvoke(new Action(BringUiToForeground), DispatcherPriority.Send);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine("[App] Activate listener invoke: " + ex.Message);
                        }
                    }
                })
                {
                    IsBackground = true,
                    Name = "MosExcelActivateListener"
                };
                _activateListenThread.Start();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[App] StartActivateListener: " + ex.Message);
            }
        }

        static void StopActivateListener()
        {
            try { _activateListenCts?.Cancel(); } catch { /* ignore */ }
            try { _activateEvent?.Set(); } catch { /* ignore */ }
            try
            {
                if (_activateListenThread != null && _activateListenThread.IsAlive)
                    _activateListenThread.Join(1000);
            }
            catch { /* ignore */ }
            _activateListenThread = null;
            try { _activateListenCts?.Dispose(); } catch { /* ignore */ }
            _activateListenCts = null;
            try { _activateEvent?.Dispose(); } catch { /* ignore */ }
            _activateEvent = null;
        }

        /// <summary>
        /// 可視ウィンドウがあれば前面化。無ければ Hidden の Main を Show して復帰する。
        /// </summary>
        void BringUiToForeground()
        {
            try
            {
                Window target = FindPreferredVisibleWindow();
                if (target == null)
                {
                    target = MainWindow
                        ?? Windows.OfType<MainWindow>().FirstOrDefault();
                    if (target == null)
                        return;
                    target.Show();
                }

                if (target.WindowState == WindowState.Minimized)
                    target.WindowState = WindowState.Normal;

                target.Show();
                target.Activate();

                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = new WindowInteropHelper(target).Handle; } catch { /* ignore */ }
                if (hwnd != IntPtr.Zero)
                {
                    ShowWindow(hwnd, SwRestore);
                    SetForegroundWindow(hwnd);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[App] BringUiToForeground: " + ex.Message);
            }
        }

        Window FindPreferredVisibleWindow()
        {
            Window appBarOrOther = null;
            foreach (Window w in Windows)
            {
                if (w == null || !w.IsLoaded || !w.IsVisible)
                    continue;

                string typeName = w.GetType().Name ?? string.Empty;
                if (typeName.IndexOf("Result", StringComparison.OrdinalIgnoreCase) >= 0)
                    return w;
                if (typeName.IndexOf("AppBar", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (appBarOrOther == null)
                        appBarOrOther = w;
                    continue;
                }

                if (appBarOrOther == null)
                    appBarOrOther = w;
            }

            return appBarOrOther;
        }

        /// <summary>既に動いている同名プロセスのメインウィンドウを前面に出す。成功時 true。</summary>
        private static bool TryActivateExistingInstance()
        {
            try
            {
                Process current = Process.GetCurrentProcess();
                string name = current.ProcessName;
                foreach (Process p in Process.GetProcessesByName(name))
                {
                    try
                    {
                        if (p.Id == current.Id)
                            continue;
                        IntPtr hwnd = p.MainWindowHandle;
                        if (hwnd == IntPtr.Zero)
                            continue;
                        ShowWindow(hwnd, SwRestore);
                        SetForegroundWindow(hwnd);
                        return true;
                    }
                    catch { /* ignore */ }
                    finally
                    {
                        try { p.Dispose(); } catch { }
                    }
                }
            }
            catch { /* ignore */ }
            return false;
        }

        const int SwRestore = 9;

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
}
