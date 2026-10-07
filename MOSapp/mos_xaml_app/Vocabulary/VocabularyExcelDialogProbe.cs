using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>Excel が開いたダイアログを探し、Esc で閉じる。</summary>
    public static class VocabularyExcelDialogProbe
    {
        static readonly string[] DialogClasses = { "#32770", "bosa_sdm_XL9", "NUIDialog" };

        delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        const uint WM_CLOSE = 0x0010;
        const byte VK_ESCAPE = 0x1B;
        const uint KEYEVENTF_KEYUP = 0x0002;

        public static uint GetProcessId(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return 0;
            GetWindowThreadProcessId(hwnd, out uint pid);
            return pid;
        }

        /// <summary>Excel のプロセスで見えているダイアログ。無ければ Zero。</summary>
        public static IntPtr FindDialog(IntPtr excelHwnd)
        {
            uint pid = GetProcessId(excelHwnd);
            if (pid == 0) return IntPtr.Zero;

            IntPtr found = IntPtr.Zero;
            var sb = new StringBuilder(64);
            EnumWindows((h, _) =>
            {
                if (h == excelHwnd || !IsWindowVisible(h)) return true;
                GetWindowThreadProcessId(h, out uint p);
                if (p != pid) return true;
                sb.Clear();
                GetClassName(h, sb, sb.Capacity);
                string cls = sb.ToString();
                foreach (var c in DialogClasses)
                {
                    if (string.Equals(cls, c, StringComparison.OrdinalIgnoreCase))
                    {
                        found = h;
                        return false;
                    }
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        /// <summary>ウィンドウを前面にして Esc だけを送る（バックステージを閉じる用）。</summary>
        public static void SendEscape(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero || !IsWindow(hwnd)) return;
            try
            {
                SetForegroundWindow(hwnd);
                System.Threading.Thread.Sleep(80);
                keybd_event(VK_ESCAPE, 0, 0, UIntPtr.Zero);
                keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            }
            catch { }
        }

        /// <summary>ダイアログを前面にして Esc を送る。閉じなければ WM_CLOSE を送る。</summary>
        public static void CloseWithEscape(IntPtr dialog)
        {
            if (dialog == IntPtr.Zero || !IsWindow(dialog)) return;
            try
            {
                SetForegroundWindow(dialog);
                System.Threading.Thread.Sleep(80);
                keybd_event(VK_ESCAPE, 0, 0, UIntPtr.Zero);
                keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            }
            catch { }

            Task.Delay(500).ContinueWith(_ =>
            {
                try
                {
                    if (IsWindow(dialog) && IsWindowVisible(dialog))
                        PostMessage(dialog, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                }
                catch { }
            });
        }
    }
}
