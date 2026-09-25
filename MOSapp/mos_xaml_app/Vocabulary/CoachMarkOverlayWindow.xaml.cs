using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;

namespace MOSExcelMogiApp.Vocabulary
{
    public partial class CoachMarkOverlayWindow : Window
    {
        const int GwlExstyle = -20;
        const int WsExTransparent = 0x00000020;
        const int WsExLayered = 0x00080000;
        const int WsExNoActivate = 0x08000000;

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public event Action Dismissed;
        bool _clickThrough;
        double _dpiScaleX = 1.0;
        double _dpiScaleY = 1.0;
        RECT _excelPhysical;
        IntPtr _excelHwnd;
        IReadOnlyList<Rect> _lastHighlightScreens = Array.Empty<Rect>();

        public CoachMarkOverlayWindow()
        {
            InitializeComponent();
            ShowActivated = false;
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            RefreshDpiScale();
            ApplyClickThrough(_clickThrough);
        }

        void RefreshDpiScale()
        {
            try
            {
                var src = PresentationSource.FromVisual(this);
                if (src?.CompositionTarget != null)
                {
                    _dpiScaleX = src.CompositionTarget.TransformToDevice.M11;
                    _dpiScaleY = src.CompositionTarget.TransformToDevice.M22;
                    if (_dpiScaleX < 0.1) _dpiScaleX = 1.0;
                    if (_dpiScaleY < 0.1) _dpiScaleY = 1.0;
                    return;
                }
            }
            catch { }

            try
            {
                var dpi = VisualTreeHelper.GetDpi(this);
                _dpiScaleX = dpi.DpiScaleX;
                _dpiScaleY = dpi.DpiScaleY;
            }
            catch
            {
                _dpiScaleX = 1.0;
                _dpiScaleY = 1.0;
            }
        }

        /// <param name="highlightScreens">画面物理ピクセル座標のハイライト矩形。</param>
        /// <param name="clickThrough">true のとき Excel へクリックを透過。</param>
        public void ShowCoachMark(
            IntPtr excelHwnd,
            IReadOnlyList<Rect> highlightScreens,
            string title,
            string message,
            bool allowDismiss = true,
            bool clickThrough = false)
        {
            _clickThrough = clickThrough;
            _excelHwnd = excelHwnd;
            _lastHighlightScreens = highlightScreens ?? Array.Empty<Rect>();
            DismissButton.Visibility = allowDismiss ? Visibility.Visible : Visibility.Collapsed;

            TitleText.Text = title ?? "";
            MessageText.Text = message ?? "";
            if (!allowDismiss && !string.IsNullOrWhiteSpace(MessageText.Text)
                && MessageText.Text.IndexOf("選択", StringComparison.Ordinal) < 0)
            {
                MessageText.Text = MessageText.Text.TrimEnd() + "\n（ハイライト箇所を選択すると次へ進みます）";
            }

            // MoveWindow(物理px)は DPI と食い違いハイライトが縮小・ずれる原因になるため使わない。
            // WPF の Left/Top/Width/Height（DIP）だけで Excel に重ねる。
            PositionOverExcel(excelHwnd);
            PaintHighlightHoles(_lastHighlightScreens);

            Bubble.IsHitTestVisible = allowDismiss && !clickThrough;
            Caret.IsHitTestVisible = false;
            HoleCanvas.IsHitTestVisible = false;

            if (!IsVisible)
                Show();

            Dispatcher.BeginInvoke(new Action(() =>
            {
                RefreshDpiScale();
                PositionOverExcel(_excelHwnd);
                PaintHighlightHoles(_lastHighlightScreens);
                ApplyClickThrough(clickThrough);
            }), System.Windows.Threading.DispatcherPriority.Loaded);

            if (!clickThrough)
                Activate();
        }

        void PositionOverExcel(IntPtr excelHwnd)
        {
            if (excelHwnd != IntPtr.Zero && GetWindowRect(excelHwnd, out RECT wr))
            {
                _excelPhysical = wr;
                RefreshDpiScale();
                Left = wr.Left / _dpiScaleX;
                Top = wr.Top / _dpiScaleY;
                Width = Math.Max(100, (wr.Right - wr.Left) / _dpiScaleX);
                Height = Math.Max(100, (wr.Bottom - wr.Top) / _dpiScaleY);
            }
            else
            {
                _excelPhysical = default;
                Left = 0;
                Top = 0;
                Width = SystemParameters.PrimaryScreenWidth;
                Height = SystemParameters.PrimaryScreenHeight;
            }
        }

        /// <summary>ハイライト矩形だけ差し替え（遅延再取得用）。</summary>
        public void UpdateHighlights(IReadOnlyList<Rect> highlightScreens)
        {
            if (!IsVisible) return;
            _lastHighlightScreens = highlightScreens ?? Array.Empty<Rect>();
            RefreshDpiScale();
            PositionOverExcel(_excelHwnd);
            PaintHighlightHoles(_lastHighlightScreens);
        }

        void PaintHighlightHoles(IReadOnlyList<Rect> highlightScreens)
        {
            HoleCanvas.Children.Clear();
            var holes = (highlightScreens ?? Array.Empty<Rect>())
                .Select(PhysicalScreenToLocalDip)
                .Where(r => r.Width >= 8 && r.Height >= 8)
                .ToList();

            RootGrid.Background = Brushes.Transparent;
            if (holes.Count > 0)
                PaintDimWithHoles(holes);
            else
                AddDim(0, 0, Width, Height);

            foreach (var localHole in holes)
            {
                var border = new Rectangle
                {
                    Width = localHole.Width,
                    Height = localHole.Height,
                    Fill = Brushes.Transparent,
                    Stroke = new SolidColorBrush(Color.FromRgb(0xFF, 0xB0, 0x20)),
                    StrokeThickness = 3,
                    RadiusX = 4,
                    RadiusY = 4,
                    IsHitTestVisible = false
                };
                Canvas.SetLeft(border, localHole.X);
                Canvas.SetTop(border, localHole.Y);
                HoleCanvas.Children.Add(border);
            }

            Rect primary = holes.Count > 0
                ? holes.OrderBy(h => h.Y).First()
                : new Rect(Width * 0.5 - 80, 40, 160, 1);
            double bubbleLeft = Math.Max(12, Math.Min(Width - 380, primary.X + primary.Width / 2 - 160));
            double bubbleTop = primary.Bottom + 16;
            if (bubbleTop + 160 > Height)
                bubbleTop = Math.Max(12, primary.Y - 160);

            Bubble.Margin = new Thickness(bubbleLeft, bubbleTop, 0, 0);
            Caret.Margin = new Thickness(
                primary.X + primary.Width / 2 - 8,
                bubbleTop < primary.Y ? bubbleTop + 130 : primary.Bottom + 8,
                0, 0);
        }

        /// <summary>画面物理ピクセル矩形 → このウィンドウ内 DIP。</summary>
        Rect PhysicalScreenToLocalDip(Rect screenPhysical)
        {
            double scaleX = _dpiScaleX <= 0 ? 1 : _dpiScaleX;
            double scaleY = _dpiScaleY <= 0 ? 1 : _dpiScaleY;
            return new Rect(
                (screenPhysical.X - _excelPhysical.Left) / scaleX,
                (screenPhysical.Y - _excelPhysical.Top) / scaleY,
                screenPhysical.Width / scaleX,
                screenPhysical.Height / scaleY);
        }

        void ApplyClickThrough(bool enable)
        {
            try
            {
                var helper = new WindowInteropHelper(this);
                IntPtr hwnd = helper.Handle;
                if (hwnd == IntPtr.Zero) return;
                int ex = GetWindowLong(hwnd, GwlExstyle);
                if (enable)
                    ex |= WsExTransparent | WsExLayered | WsExNoActivate;
                else
                    ex = (ex | WsExLayered | WsExNoActivate) & ~WsExTransparent;
                SetWindowLong(hwnd, GwlExstyle, ex);
            }
            catch { }
        }

        void PaintDimWithHoles(List<Rect> holes)
        {
            if (holes.Count == 1)
            {
                var h = holes[0];
                AddDim(0, 0, Width, Math.Max(0, h.Y));
                AddDim(0, h.Bottom, Width, Math.Max(0, Height - h.Bottom));
                AddDim(0, h.Y, Math.Max(0, h.X), h.Height);
                AddDim(h.Right, h.Y, Math.Max(0, Width - h.Right), h.Height);
                return;
            }

            AddDim(0, 0, Width, Height);
            foreach (var h in holes)
            {
                var bright = new Rectangle
                {
                    Width = h.Width,
                    Height = h.Height,
                    Fill = new SolidColorBrush(Color.FromArgb(0x55, 0xFF, 0xFF, 0xFF)),
                    IsHitTestVisible = false
                };
                Canvas.SetLeft(bright, h.X);
                Canvas.SetTop(bright, h.Y);
                HoleCanvas.Children.Add(bright);
            }
        }

        void AddDim(double x, double y, double w, double h)
        {
            if (w <= 0 || h <= 0) return;
            var r = new Rectangle
            {
                Width = w,
                Height = h,
                Fill = new SolidColorBrush(Color.FromArgb(0x99, 0, 0, 0)),
                IsHitTestVisible = false
            };
            Canvas.SetLeft(r, x);
            Canvas.SetTop(r, y);
            HoleCanvas.Children.Add(r);
        }

        void DismissButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
            Dismissed?.Invoke();
        }
    }
}
