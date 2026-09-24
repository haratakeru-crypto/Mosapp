using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
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

        static Mutex _singleInstanceMutex;
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

                bool activated = TryActivateExistingInstance();
                if (!activated)
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
            }), DispatcherPriority.ApplicationIdle);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            CloseStartupSplash();
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
