using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>テーブルチュートリアル用。「正解です！」と OK。コーチマークから押せる。</summary>
    public sealed class VocabularyTutorialOkWindow : Window
    {
        readonly Button _ok;

        public event Action OkClicked;

        public VocabularyTutorialOkWindow(string explanation)
        {
            Title = "単語帳";
            WindowStyle = WindowStyle.None;
            Background = Brushes.White;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            SizeToContent = SizeToContent.WidthAndHeight;
            ResizeMode = ResizeMode.NoResize;
            Topmost = true;
            ShowInTaskbar = false;
            ShowActivated = true;

            var panel = new StackPanel { Margin = new Thickness(28, 22, 28, 22) };
            panel.Children.Add(new TextBlock
            {
                Text = "正解です！",
                FontSize = 22,
                FontWeight = FontWeights.Bold,
                HorizontalAlignment = HorizontalAlignment.Center
            });
            if (!string.IsNullOrWhiteSpace(explanation))
            {
                panel.Children.Add(new TextBlock
                {
                    Text = explanation.Trim(),
                    FontSize = 15,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x37, 0x41, 0x51)),
                    TextWrapping = TextWrapping.Wrap,
                    TextAlignment = TextAlignment.Center,
                    MaxWidth = 360,
                    Margin = new Thickness(0, 12, 0, 0)
                });
            }
            _ok = new Button
            {
                Content = "OK",
                Width = 100,
                Height = 34,
                Margin = new Thickness(0, 18, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center,
                IsDefault = true
            };
            _ok.Click += (_, __) => OkClicked?.Invoke();
            panel.Children.Add(_ok);
            Content = panel;
        }

        public IntPtr WindowHandle => new WindowInteropHelper(this).Handle;

        public Rect TryGetWindowScreenRect()
        {
            IntPtr hwnd = WindowHandle;
            if (hwnd == IntPtr.Zero || !GetWindowRect(hwnd, out RECT wr))
                return Rect.Empty;
            return new Rect(wr.Left, wr.Top, Math.Max(8, wr.Right - wr.Left), Math.Max(8, wr.Bottom - wr.Top));
        }

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }
}
