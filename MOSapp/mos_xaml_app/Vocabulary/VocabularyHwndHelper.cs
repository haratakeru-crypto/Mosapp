using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace MOSExcelMogiApp.Vocabulary
{
    public static class VocabularyHwndHelper
    {
        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        public static IntPtr FindExcelMainWindow(ExcelApp excelApp)
        {
            int pid = -1;
            try
            {
                pid = Libraries.ExcelApplicationManager.TryGetExcelProcessId(excelApp);
            }
            catch { }

            if (pid <= 0)
            {
                try
                {
                    foreach (var p in Process.GetProcessesByName("EXCEL"))
                    {
                        pid = p.Id;
                        break;
                    }
                }
                catch { }
            }

            if (pid <= 0) return IntPtr.Zero;

            IntPtr found = IntPtr.Zero;
            EnumWindows((hWnd, _) =>
            {
                GetWindowThreadProcessId(hWnd, out uint windowPid);
                if (windowPid != (uint)pid) return true;
                if (!IsWindowVisible(hWnd)) return true;
                var sb = new StringBuilder(256);
                GetClassName(hWnd, sb, sb.Capacity);
                if (sb.ToString().Contains("XLMAIN"))
                {
                    found = hWnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);

            return found;
        }
    }
}
