using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using WordApp = Microsoft.Office.Interop.Word.Application;

namespace Libraries
{
    /// <summary>
    /// 一括採点時の Word ウィンドウ配置（PowerPoint の PositionPowerPointWindow に相当）。
    /// 最大化を解除し、画面上部に小さめ固定サイズで前面表示する。
    /// </summary>
    public static class WordWindowLayoutHelper
    {
        private const int SmCxScreen = 0;
        private const int SmCyScreen = 1;
        private const int SwRestore = 9;

        // PowerPoint 一括採点時の高さ 774 を基準に、Word は幅も抑えた固定サイズ
        private const int TargetClientWidth = 1280;
        private const int TargetClientHeight = 650;
        private const int TopMargin = 40;

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

                int screenWidth = GetSystemMetrics(SmCxScreen);
                int screenHeight = GetSystemMetrics(SmCyScreen);

                int clientWidth = Math.Min(TargetClientWidth, Math.Max(800, screenWidth - 80));
                int clientHeight = Math.Min(TargetClientHeight, Math.Max(480, screenHeight - 200));

                int x = Math.Max(0, (screenWidth - clientWidth) / 2) - borderWidth / 2;
                int y = TopMargin - borderHeight / 2;
                int width = clientWidth + borderWidth;
                int height = clientHeight + borderHeight;

                MoveWindow(wordHwnd, x, y, width, height, true);
                SetForegroundWindow(wordHwnd);

                System.Diagnostics.Debug.WriteLine(
                    $"[WordWindowLayoutHelper] Positioned for batch: {width}x{height} at ({x},{y})");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordWindowLayoutHelper] Error: {ex.Message}");
            }
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
