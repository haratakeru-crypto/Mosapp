using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private static readonly string CurrentTaskFilePath = Path.Combine(Path.GetTempPath(), "mos_excel_current_task.txt");
        private static readonly string DiagnosticLogFilePath = Path.Combine(Path.GetTempPath(), "mos_excel_addin_diag.txt");
        private bool _eventHooksRegistered;
        private Timer _taskFilePollTimer;
        private int _currentTaskProjectId = -1;
        private int _currentTaskTaskId = -1;
        private int _currentTaskAttemptNo = 1;
        private bool _ignoreNextAutoLayoutChangeAfterTaskStart;
        private string _lastRangeSelectionAddress;
        private enum LayoutChangeTrigger
        {
            Other = 0,
            SheetActivate = 1,
            WindowActivate = 2
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[ExcelAddIn1] Add-in started. Log file: " + Logger.GetLogFilePath());
            WriteDiagnostic("Startup begin");
            RegisterApplicationEventHooks();
            InitializeLayoutSnapshotsForAllOpenWorkbooks();
            StartTaskFilePolling();
            WriteDiagnostic("Startup completed");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            WriteDiagnostic("Shutdown begin");
            StopTaskFilePolling();
            UnregisterApplicationEventHooks();
            System.Diagnostics.Debug.WriteLine("[ExcelAddIn1] Add-in shutdown");
            WriteDiagnostic("Shutdown completed");
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            System.Diagnostics.Debug.WriteLine("[ExcelAddIn1] CreateRibbonExtensibilityObject called");
            return new Ribbon();
        }

        private void RegisterApplicationEventHooks()
        {
            if (_eventHooksRegistered || Application == null)
            {
                WriteDiagnostic($"RegisterApplicationEventHooks skipped. Registered={_eventHooksRegistered}, ApplicationNull={Application == null}");
                return;
            }

            Application.WorkbookOpen += Application_WorkbookOpen;
            // NewWorkbook は AppEvents_Event と _Application の両方にあり曖昧になるため明示キャストする
            ((Excel.AppEvents_Event)Application).NewWorkbook += Application_NewWorkbook;
            Application.SheetChange += Application_SheetChange;
            Application.SheetActivate += Application_SheetActivate;
            Application.SheetSelectionChange += Application_SheetSelectionChange;
            Application.WindowActivate += Application_WindowActivate;
            Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            Application.WorkbookNewChart += Application_WorkbookNewChart;
            _eventHooksRegistered = true;
            WriteDiagnostic("Application event hooks registered");
        }

        private void UnregisterApplicationEventHooks()
        {
            if (!_eventHooksRegistered || Application == null) return;

            Application.WorkbookOpen -= Application_WorkbookOpen;
            ((Excel.AppEvents_Event)Application).NewWorkbook -= Application_NewWorkbook;
            Application.SheetChange -= Application_SheetChange;
            Application.SheetActivate -= Application_SheetActivate;
            Application.SheetSelectionChange -= Application_SheetSelectionChange;
            Application.WindowActivate -= Application_WindowActivate;
            Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            Application.WorkbookNewChart -= Application_WorkbookNewChart;
            _eventHooksRegistered = false;
            WriteDiagnostic("Application event hooks unregistered");
        }

        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            string name = SafeWorkbookName(workbook);
            WriteDiagnostic("WorkbookOpen: " + name);
            InitializeLayoutSnapshotsForWorkbook(workbook, readFreeze: false);
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            string name = SafeWorkbookName(workbook);
            WriteDiagnostic("WorkbookBeforeClose: " + name);
            RemoveLayoutSnapshotsForWorkbook(workbook);
        }

        private void Application_SheetChange(object sheet, Excel.Range target)
        {
            try
            {
                string sheetName = "";
                try
                {
                    dynamic s = sheet;
                    sheetName = Convert.ToString(s?.Name) ?? "";
                }
                catch { }

                string address = "";
                try { address = target?.Address[false, false] ?? ""; } catch { }
                string operationType = IsFormulaEdit(target) ? "EditCellFormula" : "EditCellValue";
                // ユーザー編集が始まったら、TaskStart 直後の自動差分スキップを解除する。
                _ignoreNextAutoLayoutChangeAfterTaskStart = false;
                Logger.LogOperation(operationType, $"{sheetName}!{NormalizeAddress(address)}");
                WriteDiagnostic($"SheetChange: {operationType} {sheetName}!{NormalizeAddress(address)}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelAddIn1] Application_SheetChange: " + ex.Message);
                WriteDiagnostic("Application_SheetChange error: " + ex.Message);
            }
        }

        private void Application_SheetSelectionChange(object sheet, Excel.Range target)
        {
            try
            {
                _lastRangeSelectionAddress = target?.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true, Type.Missing) ?? "";
                if (_lastRangeSelectionAddress.Contains("]"))
                {
                    _lastRangeSelectionAddress = _lastRangeSelectionAddress.Substring(_lastRangeSelectionAddress.IndexOf("]") + 1);
                }
            }
            catch { }
        }

        private void Application_WorkbookNewChart(Excel.Workbook Wb, Excel.Chart Ch)
        {
            try
            {
                string selectionAddress = "";
                Excel.Range selection = Application.Selection as Excel.Range;
                if (selection != null)
                {
                    selectionAddress = selection.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true, Type.Missing);
                }
                else
                {
                    selectionAddress = _lastRangeSelectionAddress;
                }

                if (string.IsNullOrEmpty(selectionAddress))
                {
                    selectionAddress = "NoSelection";
                }

                // Remove workbook name from address if present (e.g. [book.xlsx]Sheet1!$A$1 -> Sheet1!$A$1)
                if (selectionAddress.Contains("]"))
                {
                    selectionAddress = selectionAddress.Substring(selectionAddress.IndexOf("]") + 1);
                }

                Logger.LogOperation("AddChart", $"Name={Ch.Name} Selection={selectionAddress}");
                WriteDiagnostic($"AddChart: Name={Ch.Name} Selection={selectionAddress}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelAddIn1] Application_WorkbookNewChart: " + ex.Message);
                WriteDiagnostic("Application_WorkbookNewChart error: " + ex.Message);
            }
        }

        private void StartTaskFilePolling()
        {
            if (_taskFilePollTimer != null) return;
            _taskFilePollTimer = new Timer();
            _taskFilePollTimer.Interval = 500;
            _taskFilePollTimer.Tick += TaskFilePollTimer_Tick;
            _taskFilePollTimer.Start();
        }

        private void StopTaskFilePolling()
        {
            if (_taskFilePollTimer == null) return;
            _taskFilePollTimer.Stop();
            _taskFilePollTimer.Tick -= TaskFilePollTimer_Tick;
            _taskFilePollTimer.Dispose();
            _taskFilePollTimer = null;
        }

        private void TaskFilePollTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!File.Exists(CurrentTaskFilePath)) return;

                string line;
                try
                {
                    line = File.ReadAllText(CurrentTaskFilePath).Trim();
                }
                catch
                {
                    return;
                }

                if (string.IsNullOrEmpty(line)) return;
                var parts = line.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 2) return;

                if (!int.TryParse(parts[0].Trim(), out int projectId)) return;
                if (!int.TryParse(parts[1].Trim(), out int taskId)) return;
                int attemptNo = 1;
                if (parts.Length >= 3)
                {
                    int.TryParse(parts[2].Trim(), out attemptNo);
                    if (attemptNo < 1) attemptNo = 1;
                }

                if (projectId == _currentTaskProjectId && taskId == _currentTaskTaskId && attemptNo == _currentTaskAttemptNo)
                    return;

                _currentTaskProjectId = projectId;
                _currentTaskTaskId = taskId;
                _currentTaskAttemptNo = attemptNo;

                Logger.SetCurrentTaskContext(projectId, taskId, attemptNo);
                Logger.LogTaskStart(projectId, taskId, attemptNo);
                // TaskStart 直後の自動イベント（SheetActivate/WindowActivate）で出る
                // 最初のレイアウト差分だけ無視する。
                _ignoreNextAutoLayoutChangeAfterTaskStart = true;
                InitializeLayoutSnapshotsForAllOpenWorkbooks();
                WriteDiagnostic($"Task context updated: {projectId}-{taskId}-{attemptNo} (ignore next auto layout change)");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelAddIn1] TaskFilePollTimer_Tick: " + ex.Message);
                WriteDiagnostic("TaskFilePollTimer_Tick error: " + ex.Message);
            }
        }

        private static void WriteDiagnostic(string message)
        {
            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                int pid = 0;
                try { pid = System.Diagnostics.Process.GetCurrentProcess().Id; } catch { }
                File.AppendAllText(DiagnosticLogFilePath, $"[{timestamp}] [PID:{pid}] {message}{Environment.NewLine}");
            }
            catch
            {
                // Ignore diagnostic logging errors.
            }
        }

        private bool TryConsumeAutoLayoutSuppression(LayoutChangeTrigger trigger)
        {
            if (!_ignoreNextAutoLayoutChangeAfterTaskStart)
                return false;

            if (trigger != LayoutChangeTrigger.SheetActivate && trigger != LayoutChangeTrigger.WindowActivate)
                return false;

            _ignoreNextAutoLayoutChangeAfterTaskStart = false;
            return true;
        }

        private static string NormalizeAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) return "";
            return address.Replace("$", "").Trim();
        }

        private static bool IsFormulaEdit(Excel.Range target)
        {
            if (target == null) return false;
            try
            {
                object hasFormula = target.HasFormula;
                if (hasFormula is bool b) return b;
                if (hasFormula is object[,] arr)
                {
                    foreach (var item in arr)
                    {
                        if (item is bool vb && vb) return true;
                    }
                }
            }
            catch { }
            return false;
        }

        private static string SafeWorkbookName(Excel.Workbook workbook)
        {
            try { return workbook?.Name ?? "(unknown)"; }
            catch { return "(unknown)"; }
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
