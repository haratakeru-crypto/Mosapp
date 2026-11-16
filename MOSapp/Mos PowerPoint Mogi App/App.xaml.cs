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
using Libraries;

namespace MOS_PowerPoint_app
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        const string SingleInstanceMutexName = "Local\\MOS_PowerPoint_app_SingleInstance";
        const string ActivateEventName = "Local\\MOS_PowerPoint_app_Activate";
        const int ActivateSignalWaitMs = 500;

        static Mutex _singleInstanceMutex;
        static EventWaitHandle _activateEvent;
        static CancellationTokenSource _activateListenCts;
        static Thread _activateListenThread;
        Window _startupSplash;

        protected override void OnStartup(StartupEventArgs e)
        {
            // 未処理の例外をキャッチ
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

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

                // Hidden 中の Main も起こせるよう既存プロセスへ通知（Mutex は維持）
                bool signaled = SignalExistingInstanceToActivate();
                bool activated = TryActivateExistingInstance();
                if (!signaled && !activated)
                {
                    MessageBox.Show(
                        "MOS PowerPoint アプリは既に起動しています。\n起動中の場合は、しばらくお待ちください。",
                        "MOS PowerPoint",
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

            Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose;

            // PowerPoint 画面を直接表示（表紙・パスワードは使用しない）
            var mainWindow = new MainWindow();
            Application.Current.MainWindow = mainWindow;
            mainWindow.Show();

            CloseStartupSplash();

            // UI をブロックせず VSTO 準備（開発/MSI キーの排他）
            Dispatcher.BeginInvoke(new Action(() =>
            {
                VSTOInstallerHelper.StartBackgroundPrepForExam();
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
                Text = "MOS PowerPoint アプリを起動しています。\nしばらくお待ちください。",
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
                Title = "MOS PowerPoint",
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
                {
                    // リスナ未起動（既存が古いビルド等）→ このハンドルは捨てる
                    return false;
                }

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
                    Name = "MosPowerPointActivateListener"
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

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            string errorMessage = $"未処理の例外が発生しました:\n\n{e.Exception.Message}\n\nスタックトレース:\n{e.Exception.StackTrace}";

            if (e.Exception.InnerException != null)
            {
                errorMessage += $"\n\n内部例外:\n{e.Exception.InnerException.Message}";
            }

            MessageBox.Show(errorMessage, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            System.Diagnostics.Debug.WriteLine($"DispatcherUnhandledException: {e.Exception}");

            // アプリケーションを継続させる（デバッグ用）
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;
            string errorMessage = $"致命的な例外が発生しました:\n\n{(ex != null ? ex.Message : "不明なエラー")}";

            if (ex != null)
            {
                errorMessage += $"\n\nスタックトレース:\n{ex.StackTrace}";
                if (ex.InnerException != null)
                {
                    errorMessage += $"\n\n内部例外:\n{ex.InnerException.Message}";
                }
            }

            MessageBox.Show(errorMessage, "致命的なエラー", MessageBoxButton.OK, MessageBoxImage.Error);
            System.Diagnostics.Debug.WriteLine($"UnhandledException: {ex}");
        }
    }
}
