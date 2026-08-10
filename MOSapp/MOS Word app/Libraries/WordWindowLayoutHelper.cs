using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using WordApp = Microsoft.Office.Interop.Word.Application;

namespace Libraries
{
    /// <summary>
    /// Word ウィンドウ配置ヘルパー。
    /// 試験レイアウトは Excel AppBarWindow と同じ物理ピクセル基準（GetSystemMetrics + 258/1080 スケール）。
    /// </summary>
    public static class WordWindowLayoutHelper
    {
        private const int SmCxScreen = 0;
        private const int SmCyScreen = 1;
        private const int SwRestore = 9;

        // 一括採点用: 幅も抑えた固定サイズ
        private const int TargetClientWidth = 1280;
        private const int TargetClientHeight = 650;
        private const int TopMargin = 40;

        /// <summary>アプリバー高さの設計基準（1920×1080 時の物理ピクセル）。Excel と同値。</summary>
        public const int AppBarHeightBase = 258;

        /// <summary>設計画面高さ（物理ピクセル）。Excel と同値。</summary>
        public const int DesignScreenHeight = 1080;

        /// <summary>後方互換: 設計基準のアプリバー高さ。</summary>
        public const double DefaultAppBarHeight = AppBarHeightBase;

        /// <summary>物理ピクセル単位の画面幅。</summary>
        public static int PhysicalScreenWidth => GetSystemMetrics(SmCxScreen);

        /// <summary>物理ピクセル単位の画面高さ。</summary>
        public static int PhysicalScreenHeight => GetSystemMetrics(SmCyScreen);

        /// <summary>
        /// 画面高さに合わせてスケールしたアプリバー高さ（物理ピクセル）。
        /// 1080p 以外でも Excel と同じ画面占有比率を維持する。
        /// </summary>
        public static int AppBarHeightPhysical =>
            (int)Math.Round(PhysicalScreenHeight * (double)AppBarHeightBase / DesignScreenHeight);

        /// <summary>試験時の Word 高さ = 画面高さ − アプリバー高さ（物理ピクセル）。</summary>
        public static int OfficeHeightPhysical => PhysicalScreenHeight - AppBarHeightPhysical;

        [DllImport("user32.dll")]
        private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        /// <summary>
        /// 一括採点用: Word を通常サイズに戻し、画面上部中央付近に固定して前面表示する。
        /// </summary>
        public static void PositionWordForBatchScoring(WordApp wordApp = null)
        {
            try
            {
                TryNormalizeWordWindowState(wordApp);

                IntPtr wordHwnd = FindWordMainWindowHandle();
                if (wordHwnd == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine("[WordWindowLayoutHelper] Word window handle not found");
                    return;
                }

                ShowWindow(wordHwnd, SwRestore);

                GetWindowRect(wordHwnd, out RECT windowRect);
                GetClientRect(wordHwnd, out RECT clientRect);

                int borderWidth = (windowRect.right - windowRect.left) - clientRect.right;
                int borderHeight = (windowRect.bottom - windowRect.top) - clientRect.bottom;

                int screenWidth = PhysicalScreenWidth;
                int screenHeight = PhysicalScreenHeight;

                int clientWidth = Math.Min(TargetClientWidth, Math.Max(800, screenWidth - 80));
                int clientHeight = Math.Min(TargetClientHeight, Math.Max(480, screenHeight - 200));

                int x = Math.Max(0, (screenWidth - clientWidth) / 2) - borderWidth / 2;
                int y = TopMargin - borderHeight / 2;
                int width = clientWidth + borderWidth;
                int height = clientHeight + borderHeight;

                MoveWindow(wordHwnd, x, y, width, height, true);
                SetForegroundWindow(wordHwnd);
                TryBringScoreResultToFront();

                System.Diagnostics.Debug.WriteLine(
                    $"[WordWindowLayoutHelper] Positioned for batch: {width}x{height} at ({x},{y})");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordWindowLayoutHelper] Error: {ex.Message}");
            }
        }

        /// <summary>
        /// 試験用: 最大化を解除し、画面上部にアプリバー分を除いた領域へ Word を配置する。
        /// Excel と同じく GetSystemMetrics の物理ピクセルと 258/1080 スケールを使う。
        /// </summary>
        /// <param name="appBarHeight">0 以下のときスケール済み高さを使用。正の値は上書き（物理ピクセル想定）。</param>
        /// <param name="screenWidth">0 のとき GetSystemMetrics を使用。</param>
        /// <param name="screenHeight">0 のとき GetSystemMetrics を使用。</param>
        public static void PositionWordForExamMode(
            double appBarHeight = 0,
            int screenWidth = 0,
            int screenHeight = 0)
        {
            try
            {
                TryNormalizeWordWindowState(null);

                IntPtr wordHwnd = FindWordMainWindowHandle();
                if (wordHwnd == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine("[WordWindowLayoutHelper] Word window handle not found (exam layout)");
                    return;
                }

                ShowWindow(wordHwnd, SwRestore);

                GetWindowRect(wordHwnd, out RECT windowRect);
                GetClientRect(wordHwnd, out RECT clientRect);

                int borderWidth = (windowRect.right - windowRect.left) - clientRect.right;
                int borderHeight = (windowRect.bottom - windowRect.top) - clientRect.bottom;

                if (screenWidth <= 0)
                    screenWidth = PhysicalScreenWidth;
                if (screenHeight <= 0)
                    screenHeight = PhysicalScreenHeight;

                int barH = appBarHeight > 0
                    ? (int)Math.Round(appBarHeight)
                    : (int)Math.Round(screenHeight * (double)AppBarHeightBase / DesignScreenHeight);
                int wordHeight = Math.Max(1, screenHeight - barH);
                int wordX = -borderWidth / 2;
                int wordY = -borderHeight / 2;
                int wordWidth = screenWidth + borderWidth;
                int wordHeightWithBorder = wordHeight + borderHeight;

                MoveWindow(wordHwnd, wordX, wordY, wordWidth, wordHeightWithBorder, true);
                SetForegroundWindow(wordHwnd);
                TryBringScoreResultToFront();

                System.Diagnostics.Debug.WriteLine(
                    $"[WordWindowLayoutHelper] Positioned for exam: {wordWidth}x{wordHeightWithBorder} at ({wordX},{wordY})");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordWindowLayoutHelper] Exam layout error: {ex.Message}");
            }
        }

        /// <summary>Word を前面にした直後に、開いている採点結果があれば最前面へ戻す。</summary>
        private static void TryBringScoreResultToFront()
        {
            try
            {
                MOS_Word_app.Views.ScoreResultWindow.TryBringOpenToFront();
            }
            catch { }
        }

        private static void TryNormalizeWordWindowState(WordApp wordApp)
        {
            if (wordApp == null)
            {
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    return;
                }
            }

            try
            {
                if (wordApp.ActiveWindow != null)
                {
                    wordApp.ActiveWindow.WindowState =
                        Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateNormal;
                }
            }
            catch { }
        }

        private static IntPtr FindWordMainWindowHandle()
        {
            var wordProcesses = Process.GetProcessesByName("WINWORD");
            if (wordProcesses.Length == 0)
                return IntPtr.Zero;

            IntPtr wordHwnd = IntPtr.Zero;
            try
            {
                uint processId = (uint)wordProcesses[0].Id;
                int retryCount = 0;
                const int maxRetries = 20;

                while (wordHwnd == IntPtr.Zero && retryCount < maxRetries)
                {
                    EnumWindows((windowHandle, lParam) =>
                    {
                        GetWindowThreadProcessId(windowHandle, out uint windowProcessId);
                        if (windowProcessId != processId)
                            return true;

                        var className = new StringBuilder(256);
                        GetClassName(windowHandle, className, className.Capacity);
                        if (className.ToString().Contains("OpusApp"))
                        {
                            wordHwnd = windowHandle;
                            return false;
                        }
                        return true;
                    }, IntPtr.Zero);

                    if (wordHwnd == IntPtr.Zero)
                    {
                        Thread.Sleep(200);
                        retryCount++;
                    }
                }
            }
            finally
            {
                foreach (var p in wordProcesses)
                    p.Dispose();
            }

            return wordHwnd;
        }
    }
}
