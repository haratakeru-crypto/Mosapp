using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace MOS_Word_app.Views
{
    /// <summary>
    /// プロジェクト起動・切替中の準備ダイアログ。
    /// ダイアログは専用 STA スレッドで動かし（バーが止まらない）、Word 処理は呼び出し元 UI スレッドで実行する。
    /// </summary>
    public partial class PreparingWindow : Window
    {
        static int _running;
        static PreparingWindow _active;
        static Dispatcher _dialogDispatcher;
        bool _allowClose;

        static readonly IntPtr HwndTopmost = new IntPtr(-1);

        const uint SwpNosize = 0x0001;
        const uint SwpNomove = 0x0002;
        const uint SwpShowwindow = 0x0040;

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        public PreparingWindow()
        {
            InitializeComponent();
        }

        void Window_Loaded(object sender, RoutedEventArgs e)
        {
            StartMarqueeAnimation();
        }

        void StartMarqueeAnimation()
        {
            if (MarqueeTransform == null)
                return;

            var animation = new DoubleAnimation
            {
                From = -110,
                To = 420,
                Duration = TimeSpan.FromSeconds(1.15),
                RepeatBehavior = RepeatBehavior.Forever
            };
            MarqueeTransform.BeginAnimation(TranslateTransform.XProperty, animation);
        }

        /// <summary>
        /// 準備ダイアログを専用スレッドで表示し、<paramref name="work"/> は呼び出し元スレッドで実行する。
        /// </summary>
        public static void Run(Window blockInputOn, Action work)
        {
            if (work == null)
                return;

            // すでに準備中ならダイアログは重ねず、処理だけ続ける
            if (Interlocked.CompareExchange(ref _running, 1, 0) != 0)
            {
                work();
                return;
            }

            var callerDispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;
            if (!callerDispatcher.CheckAccess())
            {
                callerDispatcher.Invoke(() => Run(blockInputOn, work));
                return;
            }

            bool restoredTopmost = false;
            bool previousTopmost = false;
            if (blockInputOn != null)
            {
                try
                {
                    previousTopmost = blockInputOn.Topmost;
                    blockInputOn.Topmost = false;
                    restoredTopmost = true;
                    blockInputOn.IsEnabled = false;
                }
                catch { /* ignore */ }
            }

            using (var ready = new ManualResetEventSlim(false))
            using (var finished = new ManualResetEventSlim(false))
            {
                Exception dialogStartError = null;

                var dialogThread = new Thread(() =>
                {
                    try
                    {
                        var dialog = new PreparingWindow();
                        _active = dialog;
                        _dialogDispatcher = Dispatcher.CurrentDispatcher;

                        dialog.Show();
                        dialog.Activate();
                        ForceTopmost(dialog);
                        FlushUntilRendered(dialog);
                        ForceTopmost(dialog);
                        ready.Set();

                        Dispatcher.Run();
                    }
                    catch (Exception ex)
                    {
                        dialogStartError = ex;
                        try { ready.Set(); } catch { /* ignore */ }
                    }
                    finally
                    {
                        _active = null;
                        _dialogDispatcher = null;
                        try { finished.Set(); } catch { /* ignore */ }
                    }
                });

                dialogThread.IsBackground = true;
                dialogThread.SetApartmentState(ApartmentState.STA);
                dialogThread.Name = "PreparingWindow";
                dialogThread.Start();

                if (!ready.Wait(8000))
                {
                    Interlocked.Exchange(ref _running, 0);
                    RestoreBlockInput(blockInputOn, restoredTopmost, previousTopmost);
                    throw new TimeoutException("準備中ダイアログの表示に失敗しました。");
                }

                if (dialogStartError != null)
                {
                    Interlocked.Exchange(ref _running, 0);
                    RestoreBlockInput(blockInputOn, restoredTopmost, previousTopmost);
                    throw new InvalidOperationException(dialogStartError.Message, dialogStartError);
                }

                try
                {
                    // Word COM などは呼び出し元 UI スレッドで実行（速さ優先）
                    work();
                }
                finally
                {
                    try
                    {
                        var d = _dialogDispatcher;
                        if (d != null && !d.HasShutdownStarted)
                        {
                            d.Invoke(() =>
                            {
                                try
                                {
                                    if (_active != null)
                                        _active.ForceClose();
                                }
                                catch { /* ignore */ }

                                try
                                {
                                    d.InvokeShutdown();
                                }
                                catch { /* ignore */ }
                            });
                        }
                    }
                    catch { /* ignore */ }

                    finished.Wait(5000);
                    RestoreBlockInput(blockInputOn, restoredTopmost, previousTopmost);
                    Interlocked.Exchange(ref _running, 0);
                }
            }
        }

        static void RestoreBlockInput(Window blockInputOn, bool restoredTopmost, bool previousTopmost)
        {
            if (blockInputOn == null)
                return;
            try { blockInputOn.IsEnabled = true; } catch { /* ignore */ }
            if (restoredTopmost)
            {
                try { blockInputOn.Topmost = previousTopmost; } catch { /* ignore */ }
            }
        }

        /// <summary>表示中の準備ダイアログを再前面化する（Word を隠した直後など）。</summary>
        public static void BringActiveToFront()
        {
            var dialog = _active;
            var dispatcher = _dialogDispatcher;
            if (dialog == null || dispatcher == null || dispatcher.HasShutdownStarted)
                return;

            try
            {
                dispatcher.BeginInvoke(new Action(() => ForceTopmost(dialog)));
            }
            catch
            {
                // ignore
            }
        }

        /// <summary>呼び出し元 UI スレッドで実行。</summary>
        public static void OnUi(Action action)
        {
            if (action == null)
                return;

            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
            {
                action();
                return;
            }

            dispatcher.Invoke(action);
        }

        static void ForceTopmost(Window dialog)
        {
            if (dialog == null)
                return;
            try
            {
                dialog.Topmost = false;
                dialog.Topmost = true;
                dialog.Activate();
                IntPtr hwnd = new WindowInteropHelper(dialog).Handle;
                if (hwnd != IntPtr.Zero)
                    SetWindowPos(hwnd, HwndTopmost, 0, 0, 0, 0, SwpNomove | SwpNosize | SwpShowwindow);
            }
            catch
            {
                // ignore
            }
        }

        static void FlushUntilRendered(Window dialog)
        {
            dialog.UpdateLayout();
            var frame = new DispatcherFrame();
            dialog.Dispatcher.BeginInvoke(DispatcherPriority.Render, new DispatcherOperationCallback(_ =>
            {
                frame.Continue = false;
                return null;
            }), null);
            Dispatcher.PushFrame(frame);
        }

        void ForceClose()
        {
            _allowClose = true;
            try { Close(); } catch { /* ignore */ }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (!_allowClose)
                e.Cancel = true;
            base.OnClosing(e);
        }
    }
}
