using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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

        const int SmXVirtualScreen = 76;
        const int SmYVirtualScreen = 77;
        const int SmCxVirtualScreen = 78;
        const int SmCyVirtualScreen = 79;
        static readonly IntPtr HwndTopmost = new IntPtr(-1);
        const uint SwpNomove = 0x0002;
        const uint SwpNosize = 0x0001;
        const uint SwpNoactivate = 0x0010;
        const uint SwpShowwindow = 0x0040;

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public event Action Dismissed;
        /// <summary>校正用: オーバーレイ上クリックの画面物理ピクセル座標。</summary>
        public event Action<Point> PhysicalClickCaptured;

        bool _clickThrough;
        bool _captureClicks;
        bool _coverScreen;
        bool _singleClick;
        Rect _bubbleScreen;
        Rect _bubbleAboveScreen;
        Rect _pinAboveScreen;
        bool _centerBubble;
        /// <summary>正解の吹き出しを、基準文面（テーブル）と同じ左上に置く。</summary>
        bool _alignToAnchor;
        string _alignAnchorMessage;
        bool _showOkButton;
        Action _singleClickAction;
        double _dpiScaleX = 1.0;
        double _dpiScaleY = 1.0;
        RECT _excelPhysical;
        IntPtr _excelHwnd;
        IReadOnlyList<Rect> _lastHighlightScreens = Array.Empty<Rect>();
        readonly List<Point> _markerLocals = new List<Point>();
        bool _dragging;
        bool _applyingPlacement;
        Point _dragStart;
        double _dragLeft;
        double _dragTop;

        public CoachMarkOverlayWindow()
        {
            InitializeComponent();
            ShowActivated = false;
            CaptureLayer.MouseLeftButtonDown += CaptureLayer_MouseLeftButtonDown;
            Bubble.MouseLeftButtonDown += Bubble_MouseLeftButtonDown;
            Bubble.MouseMove += Bubble_MouseMove;
            Bubble.MouseLeftButtonUp += Bubble_MouseLeftButtonUp;
            EditLayer.MouseLeftButtonDown += EditLayer_MouseLeftButtonDown;
            EditLayer.MouseMove += EditLayer_MouseMove;
            EditLayer.MouseLeftButtonUp += EditLayer_MouseLeftButtonUp;
            SizeChanged += (_, __) =>
            {
                if (_dragging || _applyingPlacement) return;
                PaintHighlightHoles(_lastHighlightScreens);
            };
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            RefreshDpiScale();
            ApplyClickThrough(_clickThrough);
            var src = PresentationSource.FromVisual(this) as HwndSource;
            src?.AddHook(WndProc);
        }

        /// <summary>吹き出し以外のクリックは下のウィンドウへ通す。吹き出しはドラッグできる。</summary>
        IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int wmNcHitTest = 0x0084;
            const int htTransparent = -1;
            if (msg != wmNcHitTest || _captureClicks || _singleClick || _editMode)
                return IntPtr.Zero;

            int packed = lParam.ToInt32();
            int sx = (short)(packed & 0xFFFF);
            int sy = (short)((packed >> 16) & 0xFFFF);
            Point local;
            try { local = PointFromScreen(new Point(sx, sy)); }
            catch { return IntPtr.Zero; }

            double bw = Bubble.ActualWidth > 8 ? Bubble.ActualWidth : 320;
            double bh = Bubble.ActualHeight > 8 ? Bubble.ActualHeight : 120;
            var box = new Rect(Bubble.Margin.Left, Bubble.Margin.Top, bw, bh);
            box.Inflate(12, 12);
            if (!box.Contains(local))
            {
                handled = true;
                return new IntPtr(htTransparent);
            }
            return IntPtr.Zero;
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
        /// <param name="appendSelectHint">メッセージ末尾に選択ヒントを付けるか。</param>
        public void ShowCoachMark(
            IntPtr excelHwnd,
            IReadOnlyList<Rect> highlightScreens,
            string title,
            string message,
            bool allowDismiss = true,
            bool clickThrough = false,
            bool appendSelectHint = true,
            bool coverScreen = false,
            Rect bubbleScreen = default,
            Rect bubbleAboveScreen = default,
            bool centerBubble = false,
            Rect pinAboveScreen = default,
            bool showOkButton = false,
            string alignAnchorMessage = null)
        {
            _clickThrough = clickThrough;
            _coverScreen = coverScreen;
            _bubbleScreen = bubbleScreen;
            _bubbleAboveScreen = bubbleAboveScreen;
            _pinAboveScreen = pinAboveScreen;
            _centerBubble = centerBubble;
            _alignAnchorMessage = alignAnchorMessage;
            _alignToAnchor = !string.IsNullOrWhiteSpace(alignAnchorMessage);
            _showOkButton = showOkButton;
            _excelHwnd = excelHwnd;
            _lastHighlightScreens = highlightScreens ?? Array.Empty<Rect>();
            DismissButton.Content = showOkButton ? "OK" : "閉じる";
            DismissButton.Visibility = (allowDismiss || showOkButton) ? Visibility.Visible : Visibility.Collapsed;

            TitleText.Text = title ?? "";
            MessageText.Text = message ?? "";
            if (appendSelectHint && !allowDismiss && !string.IsNullOrWhiteSpace(MessageText.Text)
                && MessageText.Text.IndexOf("選択", StringComparison.Ordinal) < 0
                && MessageText.Text.IndexOf("クリック", StringComparison.Ordinal) < 0)
            {
                MessageText.Text = MessageText.Text.TrimEnd() + "\n（ハイライト箇所を選択すると次へ進みます）";
            }
            bool placementLocked = CoachBubblePlacementStore.IsLocked(MessageText.Text);
            Bubble.Cursor = placementLocked ? Cursors.Arrow : Cursors.SizeAll;
            Bubble.ToolTip = placementLocked ? null : "ドラッグで位置を動かせます";
            PositionOverlay();
            PaintHighlightHoles(_lastHighlightScreens);

            Bubble.IsHitTestVisible = !_captureClicks;
            Caret.IsHitTestVisible = false;
            HoleCanvas.IsHitTestVisible = false;

            if (!IsVisible)
                Show();
            RaiseAboveAppBar();
            if (coverScreen)
            {
                var keepTop = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
                int raises = 0;
                keepTop.Tick += (s, e) =>
                {
                    raises++;
                    if (!IsVisible || raises > 8)
                    {
                        keepTop.Stop();
                        return;
                    }
                    RaiseAboveAppBar();
                };
                keepTop.Start();
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                RefreshDpiScale();
                PositionOverlay();
                PaintHighlightHoles(_lastHighlightScreens);
                RaiseAboveAppBar();
                ApplyClickThrough(_clickThrough && !_captureClicks && !_singleClick);
            }), System.Windows.Threading.DispatcherPriority.Loaded);

            if (!clickThrough || _captureClicks)
                Activate();
        }

        void PositionOverlay()
        {
            if (_coverScreen)
                PositionOverVirtualScreen();
            else
                PositionOverExcel(_excelHwnd);
        }

        void PositionOverVirtualScreen()
        {
            int x = GetSystemMetrics(SmXVirtualScreen);
            int y = GetSystemMetrics(SmYVirtualScreen);
            int w = Math.Max(100, GetSystemMetrics(SmCxVirtualScreen));
            int h = Math.Max(100, GetSystemMetrics(SmCyVirtualScreen));
            _excelPhysical = new RECT { Left = x, Top = y, Right = x + w, Bottom = y + h };
            RefreshDpiScale();
            Left = x / _dpiScaleX;
            Top = y / _dpiScaleY;
            Width = Math.Max(100, w / _dpiScaleX);
            Height = Math.Max(100, h / _dpiScaleY);
        }

        void RaiseAboveAppBar()
        {
            try
            {
                IntPtr hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd == IntPtr.Zero) return;
                SetWindowPos(hwnd, HwndTopmost, 0, 0, 0, 0, SwpNomove | SwpNosize | SwpNoactivate | SwpShowwindow);
            }
            catch { }
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

        public void UpdateHighlights(IReadOnlyList<Rect> highlightScreens)
        {
            if (!IsVisible) return;
            if (_editMode)
            {
                // 編集中は、まだ枠が無いときだけ後から取れた位置を入れる（手で合わせた枠を消さない）。
                if (_lastHighlightScreens != null && _lastHighlightScreens.Count > 0) return;
                _lastHighlightScreens = new List<Rect>(highlightScreens ?? Array.Empty<Rect>());
                PaintHighlightHoles(_lastHighlightScreens);
                return;
            }
            _lastHighlightScreens = highlightScreens ?? Array.Empty<Rect>();
            RefreshDpiScale();
            PositionOverlay();
            PaintHighlightHoles(_lastHighlightScreens);
        }

        public void UpdateMessage(string title, string message)
        {
            if (!string.IsNullOrEmpty(title)) TitleText.Text = title;
            if (message != null) MessageText.Text = message;
        }

        /// <summary>左上・右下クリック校正を開始。完了まで Excel へクリックを通さない。</summary>
        /// <summary>オーバーレイ上の1クリックで進む（キーワード案内など）。</summary>
        public void BeginSingleClick(Action onClick)
        {
            _singleClick = true;
            _singleClickAction = onClick;
            _captureClicks = false;
            CaptureLayer.Visibility = Visibility.Visible;
            CaptureLayer.IsHitTestVisible = true;
            ApplyClickThrough(false);
            Activate();
        }

        public void BeginClickCapture()
        {
            _singleClick = false;
            _captureClicks = true;
            _markerLocals.Clear();
            CaptureLayer.Visibility = Visibility.Visible;
            CaptureLayer.IsHitTestVisible = true;
            Bubble.IsHitTestVisible = false;
            ApplyClickThrough(false);
            Activate();
        }

        public void EndClickCapture()
        {
            _captureClicks = false;
            CaptureLayer.Visibility = Visibility.Collapsed;
            CaptureLayer.IsHitTestVisible = false;
            ApplyClickThrough(_clickThrough);
        }

        /// <summary>クリック校正と同じ基準の Excel ウィンドウ物理矩形。</summary>
        public Rect? TryGetOverlayWindowPhysical()
        {
            RefreshDpiScale();
            PositionOverExcel(_excelHwnd);
            if (_excelPhysical.Right <= _excelPhysical.Left || _excelPhysical.Bottom <= _excelPhysical.Top)
                return null;
            return new Rect(
                _excelPhysical.Left,
                _excelPhysical.Top,
                _excelPhysical.Right - _excelPhysical.Left,
                _excelPhysical.Bottom - _excelPhysical.Top);
        }

        public void ClearCalibrationMarkers()
        {
            _markerLocals.Clear();
            PaintHighlightHoles(_lastHighlightScreens);
        }

        /// <summary>校正中のクリック位置マーカー（ローカル DIP）。</summary>
        public void AddCalibrationMarkerLocal(Point localDip)
        {
            _markerLocals.Add(localDip);
            PaintHighlightHoles(_lastHighlightScreens);
        }

        void CaptureLayer_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (_singleClick)
            {
                if (!IsClickInsideHole(e.GetPosition(this)))
                {
                    e.Handled = true;
                    return;
                }
                var action = _singleClickAction;
                _singleClick = false;
                _singleClickAction = null;
                CaptureLayer.Visibility = Visibility.Collapsed;
                CaptureLayer.IsHitTestVisible = false;
                action?.Invoke();
                e.Handled = true;
                return;
            }

            if (!_captureClicks) return;
            RefreshDpiScale();
            PositionOverExcel(_excelHwnd);

            Point local = e.GetPosition(this);
            double scaleX = _dpiScaleX <= 0 ? 1 : _dpiScaleX;
            double scaleY = _dpiScaleY <= 0 ? 1 : _dpiScaleY;
            var physical = new Point(
                _excelPhysical.Left + local.X * scaleX,
                _excelPhysical.Top + local.Y * scaleY);

            AddCalibrationMarkerLocal(local);
            PhysicalClickCaptured?.Invoke(physical);
            e.Handled = true;
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

            foreach (var m in _markerLocals)
            {
                var mark = new Ellipse
                {
                    Width = 14,
                    Height = 14,
                    Fill = new SolidColorBrush(Color.FromRgb(0xFF, 0xB0, 0x20)),
                    Stroke = Brushes.White,
                    StrokeThickness = 2,
                    IsHitTestVisible = false
                };
                Canvas.SetLeft(mark, m.X - 7);
                Canvas.SetTop(mark, m.Y - 7);
                HoleCanvas.Children.Add(mark);
            }

            if (_editMode)
            {
                foreach (var localHole in holes)
                {
                    foreach (var corner in Corners(localHole))
                    {
                        var handle = new Rectangle
                        {
                            Width = HandleSize,
                            Height = HandleSize,
                            Fill = Brushes.White,
                            Stroke = new SolidColorBrush(Color.FromRgb(0xFF, 0xB0, 0x20)),
                            StrokeThickness = 2,
                            IsHitTestVisible = false
                        };
                        Canvas.SetLeft(handle, corner.X - HandleSize / 2);
                        Canvas.SetTop(handle, corner.Y - HandleSize / 2);
                        HoleCanvas.Children.Add(handle);
                    }
                }
                // 吹き出しを動かしたあと、または枠を動かしている間は、吹き出しを自動で置き直さない。
                if (_editBubbleMoved || _editHoleIndex >= 0)
                    return;
            }

            Rect primary = holes.Count > 0
                ? holes.OrderByDescending(h => h.Y).First()
                : new Rect(Width * 0.5 - 80, 40, 160, 1);

            Bubble.Measure(new Size(340, 2000));
            double bubbleHeight = Bubble.DesiredSize.Height;
            double bubbleWidth = Bubble.DesiredSize.Width;
            if (double.IsNaN(bubbleHeight) || bubbleHeight < 96) bubbleHeight = 140;
            if (double.IsNaN(bubbleWidth) || bubbleWidth < 220) bubbleWidth = 320;
            bubbleWidth = Math.Min(340, bubbleWidth);

            Rect bubbleHost = Rect.Empty;
            if (_bubbleScreen.Width >= 8 && _bubbleScreen.Height >= 8)
                bubbleHost = PhysicalScreenToLocalDip(_bubbleScreen);
            Rect aboveHost = Rect.Empty;
            if (_bubbleAboveScreen.Width >= 8 && _bubbleAboveScreen.Height >= 8)
                aboveHost = PhysicalScreenToLocalDip(_bubbleAboveScreen);
            Rect pinHost = Rect.Empty;
            if (_pinAboveScreen.Width >= 8 && _pinAboveScreen.Height >= 8)
                pinHost = PhysicalScreenToLocalDip(_pinAboveScreen);

            const double gap = 28;
            double bubbleLeft;
            double bubbleTop;
            bool caretVisible = true;
            if (pinHost.Height >= 8)
            {
                // 描画した穴（オレンジ枠）のすぐ上。ピン座標が上端に飛んでも、枠から離さない。
                Rect anchor = holes.Count > 0 ? primary : pinHost;
                if (pinHost.Height >= 8
                    && pinHost.Y >= anchor.Y - 80
                    && pinHost.Y <= anchor.Bottom + 24)
                    anchor = pinHost;
                bubbleLeft = anchor.X + (anchor.Width - bubbleWidth) / 2.0;
                bubbleTop = anchor.Y - bubbleHeight - 8;
                caretVisible = false;
            }
            else if (_centerBubble)
            {
                bubbleLeft = (Width - bubbleWidth) / 2.0;
                bubbleTop = Math.Max(48, (Height - bubbleHeight) / 2.0);
                caretVisible = false;
            }
            else if (bubbleHost.Width >= 8 && bubbleHost.Height >= 8)
            {
                // 問題文エリアの中に吹き出しを置く。
                bubbleLeft = bubbleHost.X + (bubbleHost.Width - bubbleWidth) / 2.0;
                bubbleTop = bubbleHost.Y + Math.Max(8, (bubbleHost.Height - bubbleHeight) / 2.0);
                caretVisible = false;
            }
            else if (aboveHost.Width >= 8)
            {
                bubbleLeft = aboveHost.X + aboveHost.Width / 2.0 - bubbleWidth / 2.0;
                bubbleTop = aboveHost.Y - bubbleHeight - gap;
                if (bubbleTop < 12)
                    bubbleTop = aboveHost.Bottom + gap;
            }
            else if (holes.Count == 0)
            {
                bubbleLeft = Width * 0.5 - bubbleWidth / 2.0;
                bubbleTop = Math.Max(48, Height * 0.72);
            }
            else
            {
                bubbleLeft = primary.X + primary.Width / 2.0 - bubbleWidth / 2.0;
                // 上端のタブはすぐ下、それ以外は穴のすぐ上。画面上端へは戻さない。
                bool holeNearTop = primary.Y < Math.Max(160, Height * 0.22);
                double above = primary.Y - bubbleHeight - gap;
                bubbleTop = holeNearTop || above < 12
                    ? primary.Bottom + gap
                    : above;
            }

            double extentW = Width;
            double extentH = Height;
            if (double.IsNaN(extentW) || extentW < 100) extentW = ActualWidth;
            if (double.IsNaN(extentH) || extentH < 100) extentH = ActualHeight;
            if (double.IsNaN(extentW) || extentW < 100) extentW = SystemParameters.PrimaryScreenWidth;
            if (holes.Count > 0)
                extentH = Math.Max(double.IsNaN(extentH) ? 0 : extentH, primary.Bottom + bubbleHeight + 48);
            if (double.IsNaN(extentH) || extentH < bubbleHeight + 40)
                extentH = SystemParameters.PrimaryScreenHeight;
            if (double.IsNaN(bubbleLeft)) bubbleLeft = (extentW - bubbleWidth) / 2.0;
            if (double.IsNaN(bubbleTop)) bubbleTop = holes.Count > 0 ? primary.Y - bubbleHeight - 8 : extentH * 0.7;
            if (holes.Count > 0 && bubbleTop < extentH * 0.18 && primary.Y > extentH * 0.25)
            {
                double aboveHole = primary.Y - bubbleHeight - gap;
                bubbleTop = aboveHole >= 12 ? aboveHole : primary.Bottom + gap;
                bubbleLeft = primary.X + (primary.Width - bubbleWidth) / 2.0;
            }

            // 確定時はオーバーレイに設定した幅・高さで比率を作っている。Actual はレイアウト前に
            // 小さい値のままなので、それを使うと「問題文のキーワード」だけ上にずれる。
            double layoutW = Width >= 100 ? Width : (ActualWidth >= 50 ? ActualWidth : extentW);
            double layoutH = Height >= 100 ? Height : (ActualHeight >= 50 ? ActualHeight : extentH);
            bool placedFromSaved = false;
            if (_alignToAnchor)
            {
                PlaceAtAnchor(layoutW, layoutH, ref bubbleLeft, ref bubbleTop);
                placedFromSaved = true;
            }
            else if (!_dragging && CoachBubblePlacementStore.TryGet(MessageText.Text, layoutW, layoutH, out double savedLeft, out double savedTop))
            {
                bubbleLeft = savedLeft;
                bubbleTop = savedTop;
                placedFromSaved = true;
            }

            if (placedFromSaved)
            {
                bubbleLeft = Math.Max(0, Math.Min(Math.Max(0, layoutW - 32), bubbleLeft));
                bubbleTop = Math.Max(0, Math.Min(Math.Max(0, layoutH - 32), bubbleTop));
            }
            else
            {
                double clampW = extentW;
                double clampH = extentH;
                bubbleLeft = Math.Max(12, Math.Min(Math.Max(12, clampW - bubbleWidth - 12), bubbleLeft));
                double maxTop = Math.Max(12, clampH - bubbleHeight - 12);
                bubbleTop = Math.Max(12, Math.Min(maxTop, bubbleTop));
            }

            if (_dragging)
                return;

            _applyingPlacement = true;
            try
            {
                Bubble.Margin = new Thickness(bubbleLeft, bubbleTop, 0, 0);
            }
            finally
            {
                _applyingPlacement = false;
            }
            Caret.Visibility = caretVisible ? Visibility.Visible : Visibility.Collapsed;
            double caretTop = bubbleTop < primary.Y
                ? bubbleTop + bubbleHeight + 4
                : primary.Bottom + 6;
            if (aboveHost.Width >= 8 && bubbleTop < aboveHost.Y)
                caretTop = Math.Min(bubbleTop + bubbleHeight + 2, aboveHost.Y - 10);
            Caret.Margin = new Thickness(
                holes.Count > 0 ? primary.X + primary.Width / 2 - 8 : bubbleLeft + bubbleWidth / 2 - 8,
                caretTop,
                0, 0);
        }

        bool IsClickInsideHole(Point localDip)
        {
            if (_lastHighlightScreens == null || _lastHighlightScreens.Count == 0)
                return true;
            foreach (var screen in _lastHighlightScreens)
            {
                Rect local = PhysicalScreenToLocalDip(screen);
                if (local.Width >= 8 && local.Height >= 8 && local.Contains(localDip))
                    return true;
            }
            return false;
        }

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
            // 引数は呼び出し側との互換用。全面透過にはしない。
            ApplyClickThroughCore();
        }

        void ApplyClickThroughCore()
        {
            try
            {
                var helper = new WindowInteropHelper(this);
                IntPtr hwnd = helper.Handle;
                if (hwnd == IntPtr.Zero) return;
                int ex = GetWindowLong(hwnd, GwlExstyle);
                // 全面透過にすると吹き出しをドラッグできない。外側だけ WM_NCHITTEST で通す。
                ex = (ex | WsExLayered | WsExNoActivate) & ~WsExTransparent;
                SetWindowLong(hwnd, GwlExstyle, ex);
            }
            catch { }
        }

        bool IsInsideBubbleButton(DependencyObject src)
        {
            while (src != null)
            {
                if (src == DismissButton || src == SecondaryButton) return true;
                src = VisualTreeHelper.GetParent(src);
            }
            return false;
        }

        void Bubble_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (_captureClicks
                || (!_editMode && CoachBubblePlacementStore.IsLocked(MessageText.Text))
                || IsInsideBubbleButton(e.OriginalSource as DependencyObject))
                return;
            if (_editMode) _editBubbleMoved = true;
            _dragging = true;
            _dragStart = e.GetPosition(this);
            _dragLeft = Bubble.Margin.Left;
            _dragTop = Bubble.Margin.Top;
            Bubble.CaptureMouse();
            e.Handled = true;
        }

        void Bubble_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_dragging || !Bubble.IsMouseCaptured) return;
            Point p = e.GetPosition(this);
            double dx = p.X - _dragStart.X;
            double dy = p.Y - _dragStart.Y;
            double width = ActualWidth > 50 ? ActualWidth : Width;
            double height = ActualHeight > 50 ? ActualHeight : Height;
            double bw = Bubble.ActualWidth > 8 ? Bubble.ActualWidth : 320;
            double bh = Bubble.ActualHeight > 8 ? Bubble.ActualHeight : 120;
            double left = Math.Max(8, Math.Min(Math.Max(8, width - bw - 8), _dragLeft + dx));
            double top = Math.Max(8, Math.Min(Math.Max(8, height - bh - 8), _dragTop + dy));
            Bubble.Margin = new Thickness(left, top, 0, 0);
            e.Handled = true;
        }

        void Bubble_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_dragging) return;
            _dragging = false;
            if (Bubble.IsMouseCaptured)
                Bubble.ReleaseMouseCapture();
            e.Handled = true;
            if (_editMode) Edited?.Invoke();
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

        const double HandleSize = 12;
        bool _editMode;
        bool _editBubbleMoved;
        int _editHoleIndex = -1;
        /// <summary>-1 は枠の移動、0〜3 は左上・右上・右下・左下の角。</summary>
        int _editCorner = -1;
        Point _editStartLocal;
        Rect _editStartPhysical;

        /// <summary>設定画面で吹き出しか枠を動かした。</summary>
        public event Action Edited;

        /// <summary>表示中の文面（保存のキー）。</summary>
        public string CurrentMessage => MessageText.Text;

        /// <summary>設定の保存先。正解文はテーブルの正解と同じ位置を共有する。</summary>
        public string PlacementMessage =>
            _alignToAnchor ? _alignAnchorMessage : MessageText.Text;

        /// <summary>テーブルの正解テキストが今ある左上へ、この吹き出しも置く。</summary>
        void PlaceAtAnchor(double layoutW, double layoutH, ref double bubbleLeft, ref double bubbleTop)
        {
            if (CoachBubblePlacementStore.TryGet(_alignAnchorMessage, layoutW, layoutH, out double savedLeft, out double savedTop))
            {
                bubbleLeft = savedLeft;
                bubbleTop = savedTop;
                return;
            }

            string title = TitleText.Text;
            string message = MessageText.Text;
            TitleText.Text = "正解！";
            MessageText.Text = _alignAnchorMessage;
            Bubble.Measure(new Size(340, 2000));
            double refWidth = Bubble.DesiredSize.Width;
            double refHeight = Bubble.DesiredSize.Height;
            TitleText.Text = title;
            MessageText.Text = message;
            Bubble.Measure(new Size(340, 2000));

            if (double.IsNaN(refHeight) || refHeight < 96) refHeight = 140;
            if (double.IsNaN(refWidth) || refWidth < 220) refWidth = 320;
            refWidth = Math.Min(340, refWidth);
            double extentW = Width >= 100 ? Width : layoutW;
            double extentH = Height >= 100 ? Height : layoutH;
            bubbleLeft = (extentW - refWidth) / 2.0;
            bubbleTop = Math.Max(48, (extentH - refHeight) / 2.0);
        }

        /// <summary>
        /// 設定画面用。クリックを通さず、吹き出しのドラッグと、枠の移動・四隅での大きさ変更を受け付ける。
        /// </summary>
        public void BeginEditMode()
        {
            _editMode = true;
            _editBubbleMoved = false;
            _lastHighlightScreens = new List<Rect>(_lastHighlightScreens ?? Array.Empty<Rect>());
            EditLayer.Visibility = Visibility.Visible;
            EditLayer.IsHitTestVisible = true;
            Bubble.Cursor = Cursors.SizeAll;
            Bubble.ToolTip = "ドラッグで位置を動かせます";
            DismissButton.IsEnabled = false;
            SecondaryButton.IsEnabled = false;
            PaintHighlightHoles(_lastHighlightScreens);
        }

        /// <summary>枠が無い画面に、中央へ枠を1つ足す。</summary>
        public void AddEditHole()
        {
            if (!_editMode) return;
            double sx = _dpiScaleX <= 0 ? 1 : _dpiScaleX;
            double sy = _dpiScaleY <= 0 ? 1 : _dpiScaleY;
            double w = 160, h = 60;
            var local = new Rect(Width / 2 - w / 2, Height / 2 - h / 2, w, h);
            var list = new List<Rect>(_lastHighlightScreens ?? Array.Empty<Rect>())
            {
                new Rect(_excelPhysical.Left + local.X * sx, _excelPhysical.Top + local.Y * sy, local.Width * sx, local.Height * sy)
            };
            _lastHighlightScreens = list;
            PaintHighlightHoles(_lastHighlightScreens);
            Edited?.Invoke();
        }

        /// <summary>編集後の枠（画面の物理ピクセル）。</summary>
        public List<Rect> GetEditedHoles()
        {
            return new List<Rect>(_lastHighlightScreens ?? Array.Empty<Rect>());
        }

        /// <summary>編集後の吹き出し位置と、比率の基準になるオーバーレイの大きさ（DIP）。</summary>
        public void GetEditedBubble(out double left, out double top, out double layoutW, out double layoutH)
        {
            left = Bubble.Margin.Left;
            top = Bubble.Margin.Top;
            layoutW = Width >= 100 ? Width : ActualWidth;
            layoutH = Height >= 100 ? Height : ActualHeight;
        }

        static Point[] Corners(Rect r)
        {
            return new[] { r.TopLeft, r.TopRight, r.BottomRight, r.BottomLeft };
        }

        void EditLayer_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!_editMode) return;
            Point p = e.GetPosition(this);
            var list = _lastHighlightScreens as List<Rect>;
            if (list == null) return;

            _editHoleIndex = -1;
            _editCorner = -1;
            for (int i = list.Count - 1; i >= 0 && _editHoleIndex < 0; i--)
            {
                Rect local = PhysicalScreenToLocalDip(list[i]);
                var corners = Corners(local);
                for (int c = 0; c < corners.Length; c++)
                {
                    if (Math.Abs(p.X - corners[c].X) <= HandleSize && Math.Abs(p.Y - corners[c].Y) <= HandleSize)
                    {
                        _editHoleIndex = i;
                        _editCorner = c;
                        break;
                    }
                }
                if (_editHoleIndex < 0 && local.Contains(p))
                    _editHoleIndex = i;
            }
            if (_editHoleIndex < 0) return;

            _editStartLocal = p;
            _editStartPhysical = list[_editHoleIndex];
            EditLayer.CaptureMouse();
            e.Handled = true;
        }

        void EditLayer_MouseMove(object sender, MouseEventArgs e)
        {
            var list = _lastHighlightScreens as List<Rect>;
            if (!_editMode || _editHoleIndex < 0 || list == null || _editHoleIndex >= list.Count)
            {
                UpdateEditCursor(e.GetPosition(this));
                return;
            }

            Point p = e.GetPosition(this);
            double sx = _dpiScaleX <= 0 ? 1 : _dpiScaleX;
            double sy = _dpiScaleY <= 0 ? 1 : _dpiScaleY;
            double dx = (p.X - _editStartLocal.X) * sx;
            double dy = (p.Y - _editStartLocal.Y) * sy;
            Rect r = _editStartPhysical;
            double min = 16 * sx;

            double left = r.Left, top = r.Top, right = r.Right, bottom = r.Bottom;
            switch (_editCorner)
            {
                case 0: left = Math.Min(left + dx, right - min); top = Math.Min(top + dy, bottom - min); break;
                case 1: right = Math.Max(right + dx, left + min); top = Math.Min(top + dy, bottom - min); break;
                case 2: right = Math.Max(right + dx, left + min); bottom = Math.Max(bottom + dy, top + min); break;
                case 3: left = Math.Min(left + dx, right - min); bottom = Math.Max(bottom + dy, top + min); break;
                default: left += dx; right += dx; top += dy; bottom += dy; break;
            }
            list[_editHoleIndex] = new Rect(new Point(left, top), new Point(right, bottom));
            PaintHighlightHoles(list);
            e.Handled = true;
        }

        void EditLayer_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_editHoleIndex < 0) return;
            _editHoleIndex = -1;
            _editCorner = -1;
            if (EditLayer.IsMouseCaptured) EditLayer.ReleaseMouseCapture();
            e.Handled = true;
            Edited?.Invoke();
        }

        void UpdateEditCursor(Point p)
        {
            var list = _lastHighlightScreens;
            Cursor cursor = Cursors.Arrow;
            if (list != null)
            {
                foreach (var screen in list)
                {
                    Rect local = PhysicalScreenToLocalDip(screen);
                    var corners = Corners(local);
                    for (int c = 0; c < corners.Length; c++)
                    {
                        if (Math.Abs(p.X - corners[c].X) <= HandleSize && Math.Abs(p.Y - corners[c].Y) <= HandleSize)
                            cursor = c % 2 == 0 ? Cursors.SizeNWSE : Cursors.SizeNESW;
                    }
                    if (cursor == Cursors.Arrow && local.Contains(p))
                        cursor = Cursors.SizeAll;
                }
            }
            EditLayer.Cursor = cursor;
        }

        Action _secondaryAction;

        /// <summary>OK／閉じるボタンの文字を変える。ShowCoachMark のあとに呼ぶ。</summary>
        public void SetPrimaryButtonText(string text)
        {
            if (!string.IsNullOrEmpty(text))
                DismissButton.Content = text;
        }

        /// <summary>左側に2つ目のボタンを出す。押すと閉じて onClick を呼ぶ（Dismissed は出さない）。</summary>
        public void SetSecondaryButton(string text, Action onClick)
        {
            _secondaryAction = onClick;
            SecondaryButton.Content = text ?? "";
            SecondaryButton.Visibility = string.IsNullOrEmpty(text) ? Visibility.Collapsed : Visibility.Visible;
        }

        void SecondaryButton_Click(object sender, RoutedEventArgs e)
        {
            var action = _secondaryAction;
            _secondaryAction = null;
            Close();
            action?.Invoke();
        }
    }

    /// <summary>吹き出しをドラッグした位置を、文面ごとに覚えておく。</summary>
    static class CoachBubblePlacementStore
    {
        /// <summary>画面幅・高さに対する比率。ドラッグ済みのチュートリアル文面はここを優先し、上書きしない。</summary>
        static readonly Dictionary<string, Point> Locked = new Dictionary<string, Point>(StringComparer.Ordinal)
        {
            { "問題文に出てくるキーワードがここに表示されます！", new Point(0.3953125, 0.587037037037037) },
            { "ハイライトされたテーブルをクリックして選択してください。", new Point(0.0760046487603306, 0.755716004813478) },
            { "リボンの『テーブルデザイン』タブをクリックしてください。", new Point(0.257489669421488, 0.128760529482551) },
            { "正解です！こちらのOKボタンを押して次の問題に行きましょう！", new Point(0.41171875, 0.541546869656319) },
            { "正解したら次の問題に行きましょう！", new Point(0.798978298611111, 0.827200901812819) },
            { "それでは、問題を解いてみましょう！", new Point(0.423958333333333, 0.363027777777778) },
        };

        static readonly object Gate = new object();
        static Dictionary<string, Point> _ratios;
        static Dictionary<string, Point> _confirmed;

        public static bool IsLocked(string message)
        {
            return TryFindRatio(message, out _);
        }

        static bool TryFindRatio(string message, out Point ratio)
        {
            ratio = default(Point);
            if (string.IsNullOrWhiteSpace(message)) return false;
            EnsureConfirmed();
            if (TryMatch(_confirmed, message, out ratio)) return true;
            if (Locked.TryGetValue(message, out ratio)) return true;
            string normalized = message.Replace("\r\n", "\n").Trim();
            foreach (var pair in Locked)
            {
                if (string.Equals(pair.Key, normalized, StringComparison.Ordinal))
                {
                    ratio = pair.Value;
                    return true;
                }
                if (normalized.StartsWith(pair.Key, StringComparison.Ordinal))
                {
                    ratio = pair.Value;
                    return true;
                }
            }
            return false;
        }

        static bool TryMatch(Dictionary<string, Point> map, string message, out Point ratio)
        {
            ratio = default(Point);
            if (map == null) return false;
            if (map.TryGetValue(message, out ratio)) return true;
            string normalized = message.Replace("\r\n", "\n").Trim();
            foreach (var pair in map)
            {
                if (string.Equals(pair.Key, normalized, StringComparison.Ordinal)
                    || normalized.StartsWith(pair.Key, StringComparison.Ordinal))
                {
                    ratio = pair.Value;
                    return true;
                }
            }
            return false;
        }

        public static void Confirm(string message, double left, double top, double width, double height)
        {
            if (string.IsNullOrWhiteSpace(message) || width < 50 || height < 50) return;
            EnsureConfirmed();
            var ratio = new Point(left / width, top / height);
            lock (Gate)
            {
                _confirmed[message] = ratio;
                WriteMap(ConfirmedPath, _confirmed);
            }
        }

        /// <summary>設定画面の「元に戻す」。確定した位置を消し、固定位置か自動の位置に戻す。</summary>
        public static void Unconfirm(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return;
            EnsureConfirmed();
            lock (Gate)
            {
                if (_confirmed.Remove(message))
                    WriteMap(ConfirmedPath, _confirmed);
            }
        }

        static string ConfirmedPath => System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MOSapp",
            "CoachBubbleConfirmed.txt");

        static string FilePath => System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MOSapp",
            "CoachBubblePlacements.txt");

        public static bool TryGet(string message, double width, double height, out double left, out double top)
        {
            left = 0;
            top = 0;
            if (string.IsNullOrWhiteSpace(message) || width < 50 || height < 50)
                return false;
            Point ratio;
            if (TryFindRatio(message, out ratio))
            {
                left = ratio.X * width;
                top = ratio.Y * height;
                return true;
            }
            EnsureLoaded();
            lock (Gate)
            {
                if (_ratios == null || !_ratios.TryGetValue(message, out ratio))
                    return false;
            }
            left = ratio.X * width;
            top = ratio.Y * height;
            return true;
        }

        public static void Save(string message, double left, double top, double width, double height)
        {
            if (string.IsNullOrWhiteSpace(message) || width < 50 || height < 50 || IsLocked(message))
                return;
            EnsureLoaded();
            var ratio = new Point(left / width, top / height);
            lock (Gate)
            {
                _ratios[message] = ratio;
                try
                {
                    string dir = System.IO.Path.GetDirectoryName(FilePath);
                    if (!string.IsNullOrEmpty(dir))
                        System.IO.Directory.CreateDirectory(dir);
                    var lines = new List<string>();
                    foreach (var pair in _ratios)
                    {
                        lines.Add(string.Join("\t",
                            Uri.EscapeDataString(pair.Key),
                            pair.Value.X.ToString(System.Globalization.CultureInfo.InvariantCulture),
                            pair.Value.Y.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                    }
                    System.IO.File.WriteAllLines(FilePath, lines);
                }
                catch { }
            }
        }

        static void EnsureConfirmed()
        {
            lock (Gate)
            {
                if (_confirmed != null) return;
                _confirmed = ReadMap(ConfirmedPath);
            }
        }

        static void EnsureLoaded()
        {
            lock (Gate)
            {
                if (_ratios != null) return;
                _ratios = ReadMap(FilePath);
            }
        }

        static Dictionary<string, Point> ReadMap(string path)
        {
            var map = new Dictionary<string, Point>(StringComparer.Ordinal);
            try
            {
                if (!System.IO.File.Exists(path)) return map;
                foreach (string line in System.IO.File.ReadAllLines(path))
                {
                    string[] parts = line.Split('\t');
                    if (parts.Length != 3) continue;
                    double x, y;
                    if (!double.TryParse(parts[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out x))
                        continue;
                    if (!double.TryParse(parts[2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out y))
                        continue;
                    map[Uri.UnescapeDataString(parts[0])] = new Point(x, y);
                }
            }
            catch { }
            return map;
        }

        static void WriteMap(string path, Dictionary<string, Point> map)
        {
            try
            {
                string dir = System.IO.Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir))
                    System.IO.Directory.CreateDirectory(dir);
                var lines = new List<string>();
                foreach (var pair in map)
                {
                    lines.Add(string.Join("\t",
                        Uri.EscapeDataString(pair.Key),
                        pair.Value.X.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        pair.Value.Y.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                }
                System.IO.File.WriteAllLines(path, lines);
            }
            catch { }
        }
    }
}
