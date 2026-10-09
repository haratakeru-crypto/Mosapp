using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace MOS_Word_app.Views
{
    /// <summary>
    /// 次のプロジェクト切り替えが 300ms より長いときだけ案内を出す。
    /// </summary>
    internal static class DelayedProjectOpenNotice
    {
        const int DelayMs = 300;

        public static async Task RunAsync(Func<Task> transitionAsync)
        {
            bool completed = false;
            Window overlay = null;
            DispatcherTimer keepOnTop = null;
            Timer showTimer = null;
            var dispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;

            void ShowIfNeeded()
            {
                if (completed || overlay != null)
                    return;
                try
                {
                    overlay = CreateWindow();
                    overlay.Show();
                    overlay.Activate();
                    keepOnTop = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
                    keepOnTop.Tick += (_, __) =>
                    {
                        if (overlay == null || !overlay.IsVisible || overlay.IsActive)
                            return;
                        overlay.Topmost = false;
                        overlay.Topmost = true;
                        overlay.Activate();
                    };
                    keepOnTop.Start();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[DelayedProjectOpenNotice] " + ex.Message);
                }
            }

            showTimer = new Timer(_ => dispatcher.BeginInvoke(new Action(ShowIfNeeded)), null, DelayMs, Timeout.Infinite);
            try
            {
                await dispatcher.InvokeAsync(() => { }, DispatcherPriority.ApplicationIdle);
                await transitionAsync();
            }
            finally
            {
                completed = true;
                showTimer.Dispose();
                keepOnTop?.Stop();
                if (overlay != null)
                {
                    try { overlay.Close(); } catch { }
                }
            }
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
    }
}
