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

        // ダブルクリック編集モード→確定（実値変更なし）でも SheetChange が飛ぶケース対策
        private string _pendingDoubleClickEditKey;
        private string _pendingDoubleClickEditOldValue;
        private string _pendingDoubleClickEditOldFormula;
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
            Application.SheetBeforeDoubleClick += Application_SheetBeforeDoubleClick;
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
            Application.SheetBeforeDoubleClick -= Application_SheetBeforeDoubleClick;
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

                if (ShouldSuppressNoOpDoubleClickEdit(sheet, target))
                {
                    WriteDiagnostic($"SheetChange suppressed (no-op double click): {sheetName}!{NormalizeAddress(address)}");
                }
                else
                {
                    Logger.LogOperation(operationType, $"{sheetName}!{NormalizeAddress(address)}");
                    WriteDiagnostic($"SheetChange: {operationType} {sheetName}!{NormalizeAddress(address)}");
                }

                ClearPendingDoubleClickCapture();

                TryDetectHyperlinkChangeAfterSheetChange(sheet);
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

                // 選択変更のたびにハイパーリンク集合を照合（ダイアログ挿入後に別セルを選ぶまで SheetChange が無いケースの補足）
                TryDetectHyperlinkChangeAfterSheetChange(sheet);
            }
            catch { }
        }

        private void Application_SheetBeforeDoubleClick(object sh, Excel.Range target, ref bool cancel)
        {
            try
            {
                // Cancel はしない（操作感維持）。ただし「編集開始直前の値」を一時保存して、実値変更なしの SheetChange を抑止する。
                string key = BuildPendingEditKey(sh, target);
                if (string.IsNullOrEmpty(key))
                {
                    ClearPendingDoubleClickCapture();
                    return;
                }

                _pendingDoubleClickEditKey = key;
                _pendingDoubleClickEditOldValue = SafeRangeValueText(target);
                _pendingDoubleClickEditOldFormula = SafeRangeFormulaText(target);
            }
            catch
            {
                ClearPendingDoubleClickCapture();
            }
        }

        private bool ShouldSuppressNoOpDoubleClickEdit(object sh, Excel.Range target)
        {
            try
            {
                if (target == null) return false;
                if (string.IsNullOrEmpty(_pendingDoubleClickEditKey)) return false;

                // 複数セルの変更は対象外（貼り付けなどの検知を落とさない）
                try
                {
                    if (target.CountLarge > 1) return false;
                }
                catch { return false; }

                string keyNow = BuildPendingEditKey(sh, target);
                if (!string.Equals(keyNow, _pendingDoubleClickEditKey, StringComparison.Ordinal))
                    return false;

                string newValue = SafeRangeValueText(target);
                string newFormula = SafeRangeFormulaText(target);

                bool sameValue = string.Equals(newValue ?? "", _pendingDoubleClickEditOldValue ?? "", StringComparison.Ordinal);
                bool sameFormula = string.Equals(newFormula ?? "", _pendingDoubleClickEditOldFormula ?? "", StringComparison.Ordinal);
                return sameValue && sameFormula;
            }
            catch
            {
                return false;
            }
        }

        private string BuildPendingEditKey(object sh, Excel.Range target)
        {
            try
            {
                if (target == null) return null;

                string sheetName = "";
                try
                {
                    dynamic s = sh;
                    sheetName = Convert.ToString(s?.Name) ?? "";
                }
                catch { }

                string addr = "";
                try { addr = target.Address[false, false] ?? ""; } catch { }
                if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(addr)) return null;
                return sheetName + "!" + NormalizeAddress(addr);
            }
            catch
            {
                return null;
            }
        }

        private static string SafeRangeValueText(Excel.Range r)
        {
            try
            {
                object v = r.Value2;
                if (v == null) return "";
                // 単一セル想定。型の差（数値/文字列/日付）を吸収するため文字列化して比較する。
                return Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture) ?? "";
            }
            catch { return ""; }
        }

        private static string SafeRangeFormulaText(Excel.Range r)
        {
            try
            {
                object f = r.Formula;
                if (f == null) return "";
                return Convert.ToString(f, System.Globalization.CultureInfo.InvariantCulture) ?? "";
            }
            catch { return ""; }
        }

        private void ClearPendingDoubleClickCapture()
        {
            _pendingDoubleClickEditKey = null;
            _pendingDoubleClickEditOldValue = null;
            _pendingDoubleClickEditOldFormula = null;
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

                int prevProjectId = _currentTaskProjectId;
                int prevTaskId = _currentTaskTaskId;
                int prevAttemptNo = _currentTaskAttemptNo;

                // タスク切替直前に未確定差分を「旧タスク」文脈で確定し、同一シート継続時の取りこぼしを防ぐ。
                if (prevProjectId > 0 && prevTaskId > 0)
                {
                    FlushPendingStructureDiffsForTask(prevProjectId, prevTaskId, prevAttemptNo);
                    FlushPendingSortFilterDiffsForTask(prevProjectId, prevTaskId, prevAttemptNo);
                    FlushPendingTableStyleDiffsForTask(prevProjectId, prevTaskId, prevAttemptNo);
                    FlushPendingShapeDiffsForTask(prevProjectId, prevTaskId, prevAttemptNo);
                    FlushPendingHyperlinkDiffsForTask(prevProjectId, prevTaskId, prevAttemptNo);
                }

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
