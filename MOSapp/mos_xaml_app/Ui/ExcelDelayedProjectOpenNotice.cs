using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace Ui.ViewModels
{
    /// <summary>
    /// プロジェクト切替・タスク遷移時の準備中案内（Word / PP と同様に即時表示）。
    /// </summary>
    internal static class ExcelDelayedProjectOpenNotice
    {
        public static async Task RunAsync(Func<Task> transitionAsync)
        {
            Window overlay = null;
            DispatcherTimer keepOnTop = null;
            var dispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;

            try
            {
                await dispatcher.InvokeAsync(() =>
                {
                    overlay = CreateWindow();
                    overlay.Show();
                    overlay.Activate();
                    overlay.UpdateLayout();
                    keepOnTop = StartKeepOnTop(overlay);
                });

                // COM で UI が固まる前に一度描画する
                await dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
                await dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
                await transitionAsync();
            }
            finally
            {
                keepOnTop?.Stop();
                if (overlay != null)
                {
                    try { overlay.Close(); } catch { /* ignore */ }
                }
            }
        }

        /// <summary>
        /// 同期の NavigateToTask 用。表示して描画まで進めたあと、Dispose で閉じる。
        /// </summary>
        public static IDisposable BeginImmediate()
        {
            var dispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;
            Window overlay = null;
            DispatcherTimer keepOnTop = null;

            void Show()
            {
                overlay = CreateWindow();
                overlay.Show();
                overlay.Activate();
                overlay.UpdateLayout();
                keepOnTop = StartKeepOnTop(overlay);
                FlushUntilRendered(overlay);
            }

            if (dispatcher.CheckAccess())
                Show();
            else
                dispatcher.Invoke(Show);

            return new NoticeScope(overlay, keepOnTop);
        }

        static DispatcherTimer StartKeepOnTop(Window overlay)
        {
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            timer.Tick += (_, __) =>
            {
                if (overlay == null || !overlay.IsVisible || overlay.IsActive)
                    return;
                overlay.Topmost = false;
                overlay.Topmost = true;
                overlay.Activate();
            };
            timer.Start();
            return timer;
        }

        static void FlushUntilRendered(Window dialog)
        {
            if (dialog == null)
                return;
            dialog.UpdateLayout();
            var frame = new DispatcherFrame();
            dialog.Dispatcher.BeginInvoke(DispatcherPriority.Render, new DispatcherOperationCallback(_ =>
            {
                frame.Continue = false;
                return null;
            }), null);
            Dispatcher.PushFrame(frame);
        }

        static Window CreateWindow()
        {
            return new Window
            {
                Title = "プロジェクト準備中",
                SizeToContent = SizeToContent.WidthAndHeight,
                MinWidth = 260,
                MaxWidth = 420,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.ToolWindow,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                Topmost = true,
                Content = new StackPanel
                {
                    Margin = new Thickness(16, 14, 16, 14),
                    MaxWidth = 388,
                    Children =
                    {
                        new TextBlock
                        {
                            Text = "プロジェクトを開いています。\nしばらくお待ちください...",
                            FontSize = 14,
                            TextWrapping = TextWrapping.Wrap,
                            TextAlignment = TextAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Stretch,
                            MaxWidth = 356
                        }
                    }
                }
            };
        }

        sealed class NoticeScope : IDisposable
        {
            Window _overlay;
            DispatcherTimer _keepOnTop;

            public NoticeScope(Window overlay, DispatcherTimer keepOnTop)
            {
                _overlay = overlay;
                _keepOnTop = keepOnTop;
            }

            public void Dispose()
            {
                try { _keepOnTop?.Stop(); } catch { /* ignore */ }
                _keepOnTop = null;
                if (_overlay != null)
                {
                    try { _overlay.Close(); } catch { /* ignore */ }
                    _overlay = null;
                }
            }
        }
    }
}
