using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Libraries
{
    /// <summary>
    /// リセットの開き直しが返ったあとに、ブックとアドイン起動完了を画面を止めずに確認する。
    /// </summary>
    public static class ExcelResetReadyWatch
    {
        const int TimeoutMs = 20000;
        const int PollMs = 200;

        public static void Start(int projectId, string filePath, string path = "main")
        {
            var elapsed = Stopwatch.StartNew();
            var thread = new Thread(() => Watch(projectId, filePath, elapsed, path));
            thread.IsBackground = true;
            thread.Name = "ExcelResetReady";
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        static void Watch(int projectId, string filePath, Stopwatch elapsed, string path)
        {
            try
            {
                while (elapsed.ElapsedMilliseconds < TimeoutMs)
                {
                    int processId = NewestExcelProcessId();
                    bool addInReady = processId > 0 && ExcelVstoReadiness.IsStartupCompleted(processId);
                    bool workbookReady = processId > 0 && ActiveWorkbookMatches(filePath);
                    if (addInReady && workbookReady)
                    {
                        ResetPerfLog.Write(
                            "excel",
                            projectId,
                            "ready",
                            elapsed.ElapsedMilliseconds,
                            path,
                            "signal=vsto-startup result=ok");
                        return;
                    }

                    Thread.Sleep(PollMs);
                }

                ResetPerfLog.Write(
                    "excel",
                    projectId,
                    "ready",
                    elapsed.ElapsedMilliseconds,
                    path,
                    "signal=vsto-startup result=timeout");
            }
            catch (Exception ex)
            {
                ResetPerfLog.Write(
                    "excel",
                    projectId,
                    "ready",
                    elapsed.ElapsedMilliseconds,
                    path,
                    "result=fail " + ex.Message);
            }
        }

        static int NewestExcelProcessId()
        {
            Process[] processes = Process.GetProcessesByName("EXCEL");
            try
            {
                Process newest = null;
                DateTime newestStart = DateTime.MinValue;
                foreach (Process process in processes)
                {
                    try
                    {
                        if (process.HasExited)
                            continue;
                        DateTime started = process.StartTime;
                        if (newest == null || started > newestStart)
                        {
                            newest = process;
                            newestStart = started;
                        }
                    }
                    catch
                    {
                        /* 終了途中のプロセスは飛ばす */
                    }
                }

                return newest == null ? 0 : newest.Id;
            }
            finally
            {
                foreach (Process process in processes)
                {
                    try { process.Dispose(); } catch { }
                }
            }
        }

        static bool ActiveWorkbookMatches(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return false;

            ExcelApp excelApp = null;
            try
            {
                excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                ExcelWorkbook workbook = null;
                workbook = excelApp.ActiveWorkbook;
                string fullName = null;
                try
                {
                    if (workbook != null)
                        fullName = workbook.FullName;
                }
                finally
                {
                    if (workbook != null)
                    {
                        try { Marshal.ReleaseComObject(workbook); } catch { }
                    }
                }
                if (string.IsNullOrWhiteSpace(fullName))
                    return false;
                return string.Equals(Path.GetFullPath(fullName), Path.GetFullPath(filePath), StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
            finally
            {
                if (excelApp != null)
                {
                    try { Marshal.ReleaseComObject(excelApp); } catch { }
                }
            }
        }
    }
}
