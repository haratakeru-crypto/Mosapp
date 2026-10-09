using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace MOS_Word_app.Views
{
    /// <summary>×の Reason（詳細）を平文で出す短いダイアログ。閉じるはタイトルバーの×または Esc。</summary>
    public static class ScoreReasonWindow
    {
        public static void Show(Window owner, int taskId, string reasonText)
        {
            var text = new TextBlock
            {
                Text = reasonText ?? "",
                TextWrapping = TextWrapping.Wrap,
                FontSize = 16,
                Foreground = new SolidColorBrush(Color.FromRgb(55, 65, 81))
            };

            var panel = new StackPanel { Margin = new Thickness(24) };
            panel.Children.Add(text);

            var window = new Window
            {
                Title = taskId > 0 ? $"タスク{taskId}のReason" : "Reason",
                Content = panel,
                SizeToContent = SizeToContent.WidthAndHeight,
                Width = 440,
                MinHeight = 120,
                MaxWidth = 520,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                Topmost = true,
                ShowInTaskbar = true,
                Owner = owner
            };
            window.PreviewKeyDown += (_, e) =>
            {
                if (e.Key == Key.Escape)
                {
                    window.Close();
                    e.Handled = true;
                }
            };
            window.ShowDialog();
        }
    }
}
