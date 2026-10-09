using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>単語帳の位置設定の操作パネル。オーバーレイより手前に出し続ける。</summary>
    public sealed class VocabularySettingsPanel : Window
    {
        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        static readonly IntPtr HwndTopmost = new IntPtr(-1);
        const uint SwpNomove = 0x0002, SwpNosize = 0x0001, SwpNoactivate = 0x0010;

        readonly TextBlock _position;
        readonly TextBlock _label;
        readonly TextBlock _note;
        readonly Button _addHole;
        readonly Button _operateExcel;
        readonly DispatcherTimer _keepTop;
        bool _closingByCode;

        public event Action PreviousClicked;
        public event Action NextClicked;
        public event Action SaveClicked;
        public event Action ResetClicked;
        public event Action AddHoleClicked;
        public event Action OperateExcelClicked;
        public event Action ExitClicked;

        public VocabularySettingsPanel()
        {
            Title = "単語帳の位置設定";
            Width = 460;
            SizeToContent = SizeToContent.Height;
            ResizeMode = ResizeMode.NoResize;
            WindowStyle = WindowStyle.ToolWindow;
            Topmost = true;
            ShowInTaskbar = false;
            Background = Brushes.White;

            var root = new StackPanel { Margin = new Thickness(14, 10, 14, 12) };

            _position = new TextBlock { FontSize = 13, Foreground = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80)) };
            _label = new TextBlock { FontSize = 16, FontWeight = FontWeights.Bold, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 2, 0, 4) };
            _note = new TextBlock
            {
                FontSize = 13,
                TextWrapping = TextWrapping.Wrap,
                Foreground = new SolidColorBrush(Color.FromRgb(0x37, 0x41, 0x51)),
                Text = "テキストボックスはドラッグで移動、オレンジの枠は中をドラッグで移動・四隅で大きさを変えられます。"
            };
            root.Children.Add(_position);
            root.Children.Add(_label);
            root.Children.Add(_note);

            var row1 = new WrapPanel { Margin = new Thickness(0, 10, 0, 0) };
            row1.Children.Add(MakeButton("前へ", () => PreviousClicked?.Invoke()));
            row1.Children.Add(MakeButton("次へ", () => NextClicked?.Invoke()));
            row1.Children.Add(MakeButton("保存", () => SaveClicked?.Invoke(), primary: true));
            row1.Children.Add(MakeButton("元に戻す", () => ResetClicked?.Invoke()));
            root.Children.Add(row1);

            var row2 = new WrapPanel { Margin = new Thickness(0, 6, 0, 0) };
            _addHole = MakeButton("枠を追加", () => AddHoleClicked?.Invoke());
            _operateExcel = MakeButton("Excel を操作する", () => OperateExcelClicked?.Invoke());
            row2.Children.Add(_addHole);
            row2.Children.Add(_operateExcel);
            row2.Children.Add(MakeButton("終了", () => ExitClicked?.Invoke()));
            root.Children.Add(row2);

            Content = root;

            Loaded += (_, __) =>
            {
                var area = SystemParameters.WorkArea;
                Left = area.Right - ActualWidth - 16;
                Top = area.Bottom - ActualHeight - 16;
            };

            _keepTop = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            _keepTop.Tick += (_, __) => RaiseToTop();
            _keepTop.Start();

            Closing += (s, e) =>
            {
                if (_closingByCode) return;
                // ×で閉じたときも「終了」と同じ扱いにする。
                e.Cancel = true;
                ExitClicked?.Invoke();
            };
        }

        Button MakeButton(string text, Action onClick, bool primary = false)
        {
            var button = new Button
            {
                Content = text,
                MinWidth = 96,
                Height = 32,
                Margin = new Thickness(0, 0, 8, 0),
                Padding = new Thickness(10, 0, 10, 0),
                FontSize = 13
            };
            if (primary)
            {
                button.Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x40, 0xAF));
                button.Foreground = Brushes.White;
                button.FontWeight = FontWeights.SemiBold;
            }
            button.Click += (_, __) => onClick();
            return button;
        }

        public void SetSlot(int index, int count, string label, string note, bool canAddHole)
        {
            _position.Text = (index + 1) + " / " + count;
            _label.Text = label ?? "";
            _note.Text = string.IsNullOrWhiteSpace(note)
                ? "テキストボックスはドラッグで移動、オレンジの枠は中をドラッグで移動・四隅で大きさを変えられます。"
                : note;
            _addHole.IsEnabled = canAddHole;
        }

        public void SetOperatingExcel(bool operating)
        {
            _operateExcel.Content = operating ? "表示に戻す" : "Excel を操作する";
        }

        void RaiseToTop()
        {
            try
            {
                IntPtr hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd != IntPtr.Zero)
                    SetWindowPos(hwnd, HwndTopmost, 0, 0, 0, 0, SwpNomove | SwpNosize | SwpNoactivate);
            }
            catch { }
        }

        public void CloseByCode()
        {
            _closingByCode = true;
            try { _keepTop.Stop(); } catch { }
            try { Close(); } catch { }
        }
    }
}
