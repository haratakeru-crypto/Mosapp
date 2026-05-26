using System;
using System.Collections.Generic;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    /// <summary>
    /// PageSetup / ウィンドウ枠の変化を検知し、ExcelOperationType と一致する operationType で [Op] を追記する（第1段階）。
    /// </summary>
    public partial class ThisAddIn
    {
        private readonly Dictionary<string, SheetLayoutSnapshot> _layoutSnapshots =
            new Dictionary<string, SheetLayoutSnapshot>(StringComparer.Ordinal);

        private void InitializeLayoutSnapshotsForAllOpenWorkbooks()
        {
            if (Application == null) return;
            try
            {
                foreach (Excel.Workbook wb in Application.Workbooks)
                {
                    InitializeLayoutSnapshotsForWorkbook(wb, readFreeze: false);
                }
            }
            catch (Exception ex)
            {
                WriteDiagnostic("InitializeLayoutSnapshotsForAllOpenWorkbooks: " + ex.Message);
            }
        }

        private void InitializeLayoutSnapshotsForWorkbook(Excel.Workbook workbook, bool readFreeze)
        {
            if (workbook == null) return;
            try
            {
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    TryStoreSnapshot(ws, readFreeze, logChanges: false);
                }
            }
            catch (Exception ex)
            {
                WriteDiagnostic("InitializeLayoutSnapshotsForWorkbook: " + ex.Message);
            }
        }

        private void RemoveLayoutSnapshotsForWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null) return;
            try
            {
                string wbKey = GetWorkbookKey(workbook);
                var toRemove = new List<string>();
                foreach (var kv in _layoutSnapshots)
                {
                    if (kv.Key.StartsWith(wbKey + "\x1E", StringComparison.Ordinal))
                        toRemove.Add(kv.Key);
                }
                foreach (string k in toRemove)
                    _layoutSnapshots.Remove(k);
            }
            catch (Exception ex)
            {
                WriteDiagnostic("RemoveLayoutSnapshotsForWorkbook: " + ex.Message);
            }
        }

        private void Application_SheetActivate(object sh)
        {
            try
            {
                var ws = sh as Excel.Worksheet;
                if (ws == null) return;
                try 
                { 
                    Excel.Range sel = Application.Selection as Excel.Range;
                    if (sel != null)
                    {
                        _lastRangeSelectionAddress = sel.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true, Type.Missing);
                        if (_lastRangeSelectionAddress.Contains("]"))
                        {
                            _lastRangeSelectionAddress = _lastRangeSelectionAddress.Substring(_lastRangeSelectionAddress.IndexOf("]") + 1);
                        }
                    }
                } 
                catch { }
                TryStoreSnapshot(ws, readFreeze: true, logChanges: true, trigger: LayoutChangeTrigger.SheetActivate);
            }
            catch (Exception ex)
            {
                WriteDiagnostic("Application_SheetActivate: " + ex.Message);
            }
        }

        private void Application_SheetDeactivate(object sh)
        {
            try
            {
                var ws = sh as Excel.Worksheet;
                if (ws == null) return;

                // シートが裏に隠れる直前に、現在の状態を保存し差分があればログに記録する。
                TryStoreSnapshot(ws, readFreeze: true, logChanges: true);
            }
            catch (Exception ex)
            {
                WriteDiagnostic("Application_SheetDeactivate: " + ex.Message);
            }
        }

        private void Application_WindowActivate(Excel.Workbook wb, Excel.Window wn)
        {
            try
            {
                if (wb == null) return;
                var ws = wb.ActiveSheet as Excel.Worksheet;
                if (ws == null) return;
                try 
                { 
                    Excel.Range sel = Application.Selection as Excel.Range;
                    if (sel != null)
                    {
                        _lastRangeSelectionAddress = sel.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true, Type.Missing);
                        if (_lastRangeSelectionAddress.Contains("]"))
                        {
                            _lastRangeSelectionAddress = _lastRangeSelectionAddress.Substring(_lastRangeSelectionAddress.IndexOf("]") + 1);
                        }
                    }
                } 
                catch { }
                // WindowActivate は実運用で発火頻度が低いため、主ログ経路としては使わず保険的にスナップショットだけ更新する。
                TryStoreSnapshot(ws, readFreeze: true, logChanges: false, trigger: LayoutChangeTrigger.WindowActivate);
            }
            catch (Exception ex)
            {
                WriteDiagnostic("Application_WindowActivate: " + ex.Message);
            }
        }

        private void Application_NewWorkbook(Excel.Workbook wb)
        {
            try
            {
                InitializeLayoutSnapshotsForWorkbook(wb, readFreeze: false);
            }
            catch (Exception ex)
            {
                WriteDiagnostic("Application_NewWorkbook layout: " + ex.Message);
            }
        }

        private static string GetWorkbookKey(Excel.Workbook wb)
        {
            try
            {
                string p = wb.FullName;
                if (!string.IsNullOrEmpty(p)) return p;
            }
            catch { }
            try
            {
                return wb.Name ?? "?";
            }
            catch
            {
                return "?";
            }
        }

        private static string GetSheetKey(Excel.Worksheet ws)
        {
            string wbKey = GetWorkbookKey((Excel.Workbook)ws.Parent);
            string sn = "";
            try
            {
                sn = ws.Name ?? "?";
            }
            catch
            {
                sn = "?";
            }
            return wbKey + "\x1E" + sn;
        }

        private void TryStoreSnapshot(Excel.Worksheet ws, bool readFreeze, bool logChanges, LayoutChangeTrigger trigger = LayoutChangeTrigger.Other)
        {
            SheetLayoutSnapshot? snap = BuildSnapshot(ws, readFreeze);
            if (snap == null) return;

            string key = GetSheetKey(ws);

            if (!_layoutSnapshots.TryGetValue(key, out SheetLayoutSnapshot old))
            {
                _layoutSnapshots[key] = snap.Value;
                return;
            }

            bool sheetDiff = !snap.Value.Equals(old);

            // TaskStart 直後の自動レイアウト差分は1回だけ抑止。実際にログ対象の差分があるときだけ消費する。
            bool suppressLayoutLog = false;
            if (logChanges && sheetDiff)
                suppressLayoutLog = TryConsumeAutoLayoutSuppression(trigger);

            if (sheetDiff)
            {
                if (logChanges)
                    CompareAndLogLayoutChanges(old, snap.Value, ws, suppressLayoutLog, trigger);
                _layoutSnapshots[key] = snap.Value;
            }
        }

        private static SheetLayoutSnapshot? BuildSnapshot(Excel.Worksheet ws, bool readFreeze)
        {
            try
            {
                Excel.PageSetup ps = ws.PageSetup;
                var s = new SheetLayoutSnapshot
                {
                    PrintArea = SafeGet(() => ps.PrintArea),
                    PrintTitleRows = SafeGet(() => ps.PrintTitleRows),
                    PrintTitleColumns = SafeGet(() => ps.PrintTitleColumns),
                    Orientation = SafeGetInt(() => (int)ps.Orientation),
                    LeftMargin = SafeGetDouble(() => ps.LeftMargin),
                    RightMargin = SafeGetDouble(() => ps.RightMargin),
                    TopMargin = SafeGetDouble(() => ps.TopMargin),
                    BottomMargin = SafeGetDouble(() => ps.BottomMargin),
                    LeftHeader = SafeGet(() => ps.LeftHeader),
                    CenterHeader = SafeGet(() => ps.CenterHeader),
                    RightHeader = SafeGet(() => ps.RightHeader),
                    LeftFooter = SafeGet(() => ps.LeftFooter),
                    CenterFooter = SafeGet(() => ps.CenterFooter),
                    RightFooter = SafeGet(() => ps.RightFooter),
                    Zoom = SafeGetZoom(ps),
                    PaperSize = SafeGetInt(() => (int)ps.PaperSize),
                    FitToPagesWide = SafeGetInt(() => ps.FitToPagesWide),
                    FitToPagesTall = SafeGetInt(() => ps.FitToPagesTall),
                    BlackAndWhite = SafeGetBool(() => ps.BlackAndWhite),
                    Draft = SafeGetBool(() => ps.Draft),
                    HPageBreakCount = SafeGetHBreakCount(ws),
                    VPageBreakCount = SafeGetVBreakCount(ws),
                    TableStyleSignature = BuildTableStyleSignature(ws),
                    TableRangeSignature = BuildTableRangeSignature(ws),
                    SortFilterSignature = BuildSortFilterSignature(ws),
                    ShapeCount = SafeGetShapeCount(ws),
                    ShapeGeometrySignature = BuildShapeGeometrySignature(ws),
                    NamedRangeSignature = BuildNamedRangeSignature(ws),
                    ExternalDataSignature = BuildExternalDataSignature(ws),
                    ConditionalFormatSignature = BuildConditionalFormatSignature(ws),
                    UsedRowCount = SafeGetUsedRangeRowCount(ws),
                    UsedColumnCount = SafeGetUsedRangeColumnCount(ws),
                    CellFormatSignature = BuildCellFormatSignature(ws),
                    HyperlinkSignature = BuildHyperlinkSignature(ws)
                };

                if (readFreeze && TryGetFreezeForActiveSheet(ws, out bool freeze, out int splitRow, out int splitCol))
                {
                    s.FreezePanes = freeze;
                    s.SplitRow = splitRow;
                    s.SplitColumn = splitCol;
                }
                else
                {
                    s.FreezePanes = null;
                    s.SplitRow = null;
                    s.SplitColumn = null;
                }

                return s;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[LayoutMonitoring] BuildSnapshot: " + ex.Message);
                return null;
            }
        }

        private static bool TryGetFreezeForActiveSheet(Excel.Worksheet ws, out bool freeze, out int splitRow, out int splitCol)
        {
            freeze = false;
            splitRow = 0;
            splitCol = 0;
            try
            {
                var app = ws.Application;
                if (app.ActiveSheet is Excel.Worksheet activeWs &&
                    string.Equals(activeWs.Name, ws.Name, StringComparison.OrdinalIgnoreCase) &&
                    GetWorkbookKey((Excel.Workbook)activeWs.Parent) == GetWorkbookKey((Excel.Workbook)ws.Parent))
                {
                    Excel.Window w = app.ActiveWindow;
                    if (w != null)
                    {
                        freeze = w.FreezePanes;
                        splitRow = w.SplitRow;
                        splitCol = w.SplitColumn;
                        return true;
                    }
                }
            }
            catch { }
            return false;
        }

        private static int SafeGetHBreakCount(Excel.Worksheet ws)
        {
            try
            {
                return ws.HPageBreaks.Count;
            }
            catch
            {
                return -1;
            }
        }

        private static int SafeGetVBreakCount(Excel.Worksheet ws)
        {
            try
            {
                return ws.VPageBreaks.Count;
            }
            catch
            {
                return -1;
            }
        }

        private static string SafeGet(Func<string> f)
        {
            try
            {
                return f() ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static int SafeGetInt(Func<int> f)
        {
            try
            {
                return f();
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>ページに合わせる等の状態では Zoom が無効で例外になることがある。</summary>
        private static int SafeGetZoom(Excel.PageSetup ps)
        {
            try
            {
                object z = ps.Zoom;
                if (z == null) return 0;
                if (z is bool) return 0;
                return Convert.ToInt32(Math.Round(Convert.ToDouble(z)));
            }
            catch
            {
                return 0;
            }
        }

        private static double SafeGetDouble(Func<double> f)
        {
            try
            {
                return f();
            }
            catch
            {
                return double.NaN;
            }
        }

        private static bool SafeGetBool(Func<bool> f)
        {
            try
            {
                return f();
            }
            catch
            {
                return false;
            }
        }

        private static string BuildTableStyleSignature(Excel.Worksheet ws)
        {
            var parts = new List<string>();
            try
            {
                foreach (Excel.ListObject table in ws.ListObjects)
                {
                    string name = SafeGet(() => table.Name);
                    string style = "";
                    try
                    {
                        // TableStyle は COM object のため ToString で正規化する。
                        style = table.TableStyle == null ? "" : Convert.ToString(table.TableStyle) ?? "";
                    }
                    catch { }
                    parts.Add($"{name}:{style}");
                }
            }
            catch { }
            parts.Sort(StringComparer.Ordinal);
            return string.Join("|", parts);
        }

        private static string BuildTableRangeSignature(Excel.Worksheet ws)
        {
            var parts = new List<string>();
            try
            {
                foreach (Excel.ListObject table in ws.ListObjects)
                {
                    string name = SafeGet(() => table.Name);
                    string addr = "";
                    try
                    {
                        addr = table.Range?.Address[false, false] ?? "";
                    }
                    catch { }
                    parts.Add($"{name}:{addr}");
                }
            }
            catch { }
            parts.Sort(StringComparer.Ordinal);
            return string.Join("|", parts);
        }

        private static string BuildSortFilterSignature(Excel.Worksheet ws)
        {
            try
            {
                var parts = new List<string>();

                // Filter 署名
                try
                {
                    Excel.AutoFilter af = ws.AutoFilter;
                    if (af != null)
                    {
                        string afRange = "";
                        try { afRange = af.Range?.Address[false, false] ?? ""; } catch { }

                        parts.Add("AFR:" + afRange);
                        int count = 0;
                        try { count = af.Filters.Count; } catch { }

                        for (int i = 1; i <= count; i++)
                        {
                            try
                            {
                                var filter = af.Filters[i] as Excel.Filter;
                                if (filter == null || !filter.On) continue;
                                string c1 = SafeObjectToText(filter.Criteria1);
                                string c2 = SafeObjectToText(filter.Criteria2);
                                int op = 0;
                                try { op = (int)filter.Operator; } catch { }
                                parts.Add($"AF:{i}:{op}:{c1}:{c2}");
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        parts.Add("AFR:");
                    }
                }
                catch
                {
                    parts.Add("AFERR");
                }

                // Sort 署名（並べ替えを拾うため、Worksheet.Sort のキー・順序を含める）
                try
                {
                    Excel.Sort sort = ws.Sort;
                    if (sort != null)
                    {
                        string sortRange = "";
                        try { sortRange = sort.Rng?.Address[false, false] ?? ""; } catch { }
                        string header = "";
                        try { header = SafeObjectToText(sort.Header); } catch { }
                        string orientation = "";
                        try { orientation = SafeObjectToText(sort.Orientation); } catch { }
                        string method = "";
                        try { method = SafeObjectToText(sort.SortMethod); } catch { }
                        string matchCase = "";
                        try { matchCase = SafeObjectToText(sort.MatchCase); } catch { }

                        parts.Add($"SR:{sortRange}:H={header}:O={orientation}:M={method}:C={matchCase}");

                        Excel.SortFields fields = null;
                        try { fields = sort.SortFields; } catch { }
                        int sfCount = 0;
                        try { sfCount = fields?.Count ?? 0; } catch { }
                        for (int i = 1; i <= sfCount; i++)
                        {
                            try
                            {
                                Excel.SortField sf = fields[i];
                                string key = "";
                                try { key = sf.Key?.Address[false, false] ?? ""; } catch { }
                                string order = "";
                                try { order = SafeObjectToText(sf.Order); } catch { }
                                string sortOn = "";
                                try { sortOn = SafeObjectToText(sf.SortOn); } catch { }
                                string dataOption = "";
                                try { dataOption = SafeObjectToText(sf.DataOption); } catch { }
                                parts.Add($"SF:{i}:{key}:ON={sortOn}:ORD={order}:OPT={dataOption}");
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        parts.Add("SR:");
                    }
                }
                catch
                {
                    parts.Add("SRERR");
                }

                return string.Join("|", parts);
            }
            catch
            {
                return "";
            }
        }

        private static string SafeObjectToText(object value)
        {
            try
            {
                if (value == null) return "";
                if (value is Array arr)
                {
                    var items = new List<string>();
                    foreach (object item in arr)
                    {
                        items.Add(item?.ToString() ?? "");
                    }
                    return "[" + string.Join(",", items) + "]";
                }
                return value.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static int SafeGetShapeCount(Excel.Worksheet ws)
        {
            try
            {
                return ws.Shapes?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static string BuildShapeGeometrySignature(Excel.Worksheet ws)
        {
            var parts = new List<string>();
            try
            {
                foreach (Excel.Shape shape in ws.Shapes)
                {
                    try
                    {
                        string id = "";
                        try { id = shape.Name ?? ""; } catch { }
                        string type = "";
                        try { type = ((int)shape.Type).ToString(CultureInfo.InvariantCulture); } catch { }
                        double left = 0, top = 0, width = 0, height = 0;
                        try { left = Math.Round(shape.Left, 2); } catch { }
                        try { top = Math.Round(shape.Top, 2); } catch { }
                        try { width = Math.Round(shape.Width, 2); } catch { }
                        try { height = Math.Round(shape.Height, 2); } catch { }
                        parts.Add($"{id}:{type}:{left.ToString("0.00", CultureInfo.InvariantCulture)}:{top.ToString("0.00", CultureInfo.InvariantCulture)}:{width.ToString("0.00", CultureInfo.InvariantCulture)}:{height.ToString("0.00", CultureInfo.InvariantCulture)}");
                    }
                    catch { }
                }
            }
            catch { }

            parts.Sort(StringComparer.Ordinal);
            return string.Join("|", parts);
        }

        private static string BuildNamedRangeSignature(Excel.Worksheet ws)
        {
            var parts = new List<string>();
            try
            {
                var wb = ws.Parent as Excel.Workbook;
                if (wb != null)
                {
                    foreach (Excel.Name n in wb.Names)
                    {
                        try
                        {
                            string name = SafeGet(() => n.Name);
                            string refersTo = SafeGet(() => n.RefersTo);
                            parts.Add($"WB:{name}:{refersTo}");
                        }
                        catch { }
                    }
                }

                foreach (Excel.Name n in ws.Names)
                {
                    try
                    {
                        string name = SafeGet(() => n.Name);
                        string refersTo = SafeGet(() => n.RefersTo);
                        parts.Add($"WS:{name}:{refersTo}");
                    }
                    catch { }
                }
            }
            catch { }

            parts.Sort(StringComparer.Ordinal);
            return string.Join("|", parts);
        }

        private static string BuildExternalDataSignature(Excel.Worksheet ws)
        {
            var parts = new List<string>();
            try
            {
                foreach (Excel.QueryTable qt in ws.QueryTables)
                {
                    try
                    {
                        string name = SafeGet(() => qt.Name);
                        string conn = SafeGet(() => qt.Connection);
                        string dst = "";
                        try { dst = qt.ResultRange?.Address[false, false] ?? ""; } catch { }
                        parts.Add($"QT:{name}:{dst}:{conn}");
                    }
                    catch { }
                }

                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    try
                    {
                        Excel.QueryTable qt = null;
                        try { qt = lo.QueryTable; } catch { }
                        if (qt == null) continue;

                        string name = SafeGet(() => lo.Name);
                        string conn = SafeGet(() => qt.Connection);
                        string dst = "";
                        try { dst = lo.Range?.Address[false, false] ?? ""; } catch { }
                        parts.Add($"LOQT:{name}:{dst}:{conn}");
                    }
                    catch { }
                }
            }
            catch { }

            parts.Sort(StringComparer.Ordinal);
            return string.Join("|", parts);
        }

        private static string BuildConditionalFormatSignature(Excel.Worksheet ws)
        {
            var parts = new List<string>();
            try
            {
                Excel.Range used = null;
                try { used = ws.UsedRange; } catch { }
                if (used == null) return "";

                Excel.FormatConditions fcs = null;
                try { fcs = used.FormatConditions; } catch { }
                if (fcs == null) return "";

                int count = 0;
                try { count = fcs.Count; } catch { }
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        var fc = fcs.Item(i);
                        if (fc == null) continue;
                        string type = "";
                        string op = "";
                        string f1 = "";
                        string f2 = "";
                        try { type = SafeObjectToText(fc.Type); } catch { }
                        try { op = SafeObjectToText(fc.Operator); } catch { }
                        try { f1 = SafeObjectToText(fc.Formula1); } catch { }
                        try { f2 = SafeObjectToText(fc.Formula2); } catch { }
                        parts.Add($"{i}:{type}:{op}:{f1}:{f2}");
                    }
                    catch { }
                }
            }
            catch { }

            parts.Sort(StringComparer.Ordinal);
            return string.Join("|", parts);
        }

        private static int SafeGetUsedRangeRowCount(Excel.Worksheet ws)
        {
            try
            {
                Excel.Range used = ws.UsedRange;
                if (used == null) return 0;
                return used.Rows?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static int SafeGetUsedRangeColumnCount(Excel.Worksheet ws)
        {
            try
            {
                Excel.Range used = ws.UsedRange;
                if (used == null) return 0;
                return used.Columns?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static string BuildCellFormatSignature(Excel.Worksheet ws)
        {
            try
            {
                Excel.Range used = ws.UsedRange;
                if (used == null) return "";

                string addr = "";
                try { addr = used.Address[false, false] ?? ""; } catch { }

                string numberFormat = "";
                try { numberFormat = SafeObjectToText(used.NumberFormat); } catch { }

                string wrap = "";
                try { wrap = SafeObjectToText(used.WrapText); } catch { }

                string hAlign = "";
                try { hAlign = SafeObjectToText(used.HorizontalAlignment); } catch { }

                string vAlign = "";
                try { vAlign = SafeObjectToText(used.VerticalAlignment); } catch { }

                string style = "";
                try { style = SafeObjectToText(used.Style); } catch { }

                return $"{addr}|NF:{numberFormat}|W:{wrap}|H:{hAlign}|V:{vAlign}|S:{style}";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>シート上のハイパーリンク集合の署名（挿入・変更・削除の差分検知用）。</summary>
        private static string BuildHyperlinkSignature(Excel.Worksheet ws)
        {
            var parts = new List<string>();
            try
            {
                Excel.Hyperlinks hls = ws.Hyperlinks;
                if (hls == null) return "";

                int count = 0;
                try { count = hls.Count; } catch { }
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        Excel.Hyperlink hl = hls[i];
                        string addr = "";
                        try { addr = hl.Address ?? ""; } catch { }
                        string sub = "";
                        try { sub = hl.SubAddress ?? ""; } catch { }
                        string rng = "";
                        try { rng = hl.Range?.Address[false, false] ?? ""; } catch { }
                        parts.Add($"{rng}|{addr}|{sub}");
                    }
                    catch { }
                }
            }
            catch { }

            parts.Sort(StringComparer.Ordinal);
            return string.Join("|", parts);
        }

        /// <summary>
        /// <see cref="Application_SheetChange"/> 後にハイパーリンク集合だけ比較する。
        /// ハイパーリンク挿入直後にシート切替が無くても <c>[Op] InsertHyperlink</c> を残す。
        /// </summary>
        private void TryDetectHyperlinkChangeAfterSheetChange(object sheet)
        {
            var ws = sheet as Excel.Worksheet;
            if (ws == null) return;
            try
            {
                string key = GetSheetKey(ws);
                if (!_layoutSnapshots.TryGetValue(key, out SheetLayoutSnapshot old))
                    return;

                string newHl = BuildHyperlinkSignature(ws);
                if (old.HyperlinkSignature == newHl) return;

                string sheetName = "";
                try { sheetName = ws.Name ?? "?"; } catch { sheetName = "?"; }

                Logger.LogOperation("InsertHyperlink", $"{sheetName}!HyperlinkChanged");

                // HyperlinkSignature だけ更新すると他フィールドが古くなり次の TryStoreSnapshot で誤差分になるためフル更新する
                SheetLayoutSnapshot? fullSnap = BuildSnapshot(ws, readFreeze: false);
                if (fullSnap != null)
                    _layoutSnapshots[key] = fullSnap.Value;
                else
                {
                    SheetLayoutSnapshot updated = old;
                    updated.HyperlinkSignature = newHl;
                    _layoutSnapshots[key] = updated;
                }
            }
            catch (Exception ex)
            {
                WriteDiagnostic("TryDetectHyperlinkChangeAfterSheetChange: " + ex.Message);
            }
        }

        /// <summary>
        /// タスク切替直前の境界フラッシュ。選択シートのレイアウト（PageSetup 等）を先に確定し、
        /// 全ワークブックの構造・並べ替え/フィルター・テーブルスタイル・図形・ハイパーリンクは
        /// 各シートで <see cref="BuildSnapshot"/> を1回だけ実行してまとめて判定する。
        /// </summary>
        private void FlushPendingBoundaryDiffsForTask(int projectId, int taskId, int attemptNo)
        {
            if (projectId <= 0 || taskId <= 0) return;
            FlushPendingSelectedSheetsLayoutDiffsForTask(projectId, taskId, attemptNo);
            FlushPendingWorkbookSheetsBoundaryDiffsSinglePass(projectId, taskId, attemptNo);
        }

        /// <summary>
        /// 全ブック各シートについて、旧スナップショットと現在状態を1回の <c>BuildSnapshot</c> で比較し、
        /// 差分カテゴリをまとめてログしてからスナップショットを更新する。
        /// </summary>
        private void FlushPendingWorkbookSheetsBoundaryDiffsSinglePass(int projectId, int taskId, int attemptNo)
        {
            if (projectId <= 0 || taskId <= 0 || Application == null) return;
            try
            {
                foreach (Excel.Workbook wb in Application.Workbooks)
                {
                    try
                    {
                        foreach (Excel.Worksheet ws in wb.Worksheets)
                        {
                            try
                            {
                                string key = GetSheetKey(ws);
                                SheetLayoutSnapshot? snap = BuildSnapshot(ws, readFreeze: false);
                                if (snap == null) continue;

                                if (!_layoutSnapshots.TryGetValue(key, out SheetLayoutSnapshot old))
                                {
                                    _layoutSnapshots[key] = snap.Value;
                                    continue;
                                }

                                SheetLayoutSnapshot now = snap.Value;
                                bool rowChanged = old.UsedRowCount != now.UsedRowCount;
                                bool colChanged = old.UsedColumnCount != now.UsedColumnCount;
                                bool sortFilterChanged = old.SortFilterSignature != now.SortFilterSignature;
                                bool tableStyleChanged = old.TableStyleSignature != now.TableStyleSignature;
                                bool shapeCountChanged = old.ShapeCount != now.ShapeCount;
                                bool shapeGeomChanged = old.ShapeGeometrySignature != now.ShapeGeometrySignature;
                                bool hyperlinkChanged = old.HyperlinkSignature != now.HyperlinkSignature;

                                if (!rowChanged && !colChanged && !sortFilterChanged && !tableStyleChanged
                                    && !shapeCountChanged && !shapeGeomChanged && !hyperlinkChanged)
                                {
                                    continue;
                                }

                                string sheetName = "";
                                try { sheetName = ws.Name ?? "?"; } catch { sheetName = "?"; }

                                Logger.RunWithTaskContext(projectId, taskId, attemptNo, () =>
                                {
                                    if (rowChanged)
                                    {
                                        if (now.UsedRowCount > old.UsedRowCount)
                                            Logger.LogOperation("InsertRows", $"{sheetName}!Rows:{old.UsedRowCount}->{now.UsedRowCount};Trigger=BoundaryFlush");
                                        else
                                            Logger.LogOperation("DeleteRows", $"{sheetName}!Rows:{old.UsedRowCount}->{now.UsedRowCount};Trigger=BoundaryFlush");
                                    }

                                    if (colChanged)
                                    {
                                        if (now.UsedColumnCount > old.UsedColumnCount)
                                            Logger.LogOperation("InsertColumns", $"{sheetName}!Cols:{old.UsedColumnCount}->{now.UsedColumnCount};Trigger=BoundaryFlush");
                                        else
                                            Logger.LogOperation("DeleteColumns", $"{sheetName}!Cols:{old.UsedColumnCount}->{now.UsedColumnCount};Trigger=BoundaryFlush");
                                    }

                                    if (sortFilterChanged)
                                        Logger.LogOperation("SortOrFilter", $"{sheetName}!{now.SortFilterSignature};Trigger=BoundaryFlush");

                                    if (tableStyleChanged)
                                        Logger.LogOperation("SetTableStyle", $"{sheetName}!{now.TableStyleSignature};Trigger=BoundaryFlush");

                                    if (shapeCountChanged)
                                    {
                                        if (now.ShapeCount > old.ShapeCount)
                                            Logger.LogOperation("InsertShapeOrImage", $"{sheetName}!Count:{old.ShapeCount}->{now.ShapeCount};Trigger=BoundaryFlush");
                                        else
                                            Logger.LogOperation("DeleteShapeOrImage", $"{sheetName}!Count:{old.ShapeCount}->{now.ShapeCount};Trigger=BoundaryFlush");
                                    }
                                    else if (shapeGeomChanged)
                                    {
                                        Logger.LogOperation("MoveOrResizeShape", $"{sheetName}!Count={now.ShapeCount};Trigger=BoundaryFlush");
                                    }

                                    if (hyperlinkChanged)
                                        Logger.LogOperation("InsertHyperlink", $"{sheetName}!HyperlinkChanged;Trigger=BoundaryFlush");
                                });

                                _layoutSnapshots[key] = now;
                            }
                            catch (Exception exInner)
                            {
                                WriteDiagnostic("FlushPendingWorkbookSheetsBoundaryDiffsSinglePass sheet: " + exInner.Message);
                            }
                        }
                    }
                    catch (Exception exWb)
                    {
                        WriteDiagnostic("FlushPendingWorkbookSheetsBoundaryDiffsSinglePass workbook: " + exWb.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDiagnostic("FlushPendingWorkbookSheetsBoundaryDiffsSinglePass: " + ex.Message);
            }
        }

        /// <summary>
        /// タスク切替直前に、現在選択されているすべてのシート（ActiveSheet含む）のレイアウト差分を旧タスク文脈で確定する。
        /// 印刷設定や書式などは全シート回すと重いため、ユーザーが直前まで触っていた可能性が高い選択シートのみに限定してチェックする。
        /// </summary>
        private void FlushPendingSelectedSheetsLayoutDiffsForTask(int projectId, int taskId, int attemptNo)
        {
            if (projectId <= 0 || taskId <= 0 || Application == null) return;
            try
            {
                Excel.Window activeWindow = Application.ActiveWindow;
                if (activeWindow == null) return;

                Excel.Sheets selectedSheets = activeWindow.SelectedSheets;
                if (selectedSheets == null) return;

                foreach (object sh in selectedSheets)
                {
                    try
                    {
                        var ws = sh as Excel.Worksheet;
                        if (ws == null) continue;

                        string key = GetSheetKey(ws);
                        SheetLayoutSnapshot? snap = BuildSnapshot(ws, readFreeze: true);
                        if (snap == null) continue;

                        if (!_layoutSnapshots.TryGetValue(key, out SheetLayoutSnapshot old))
                        {
                            _layoutSnapshots[key] = snap.Value;
                            continue;
                        }

                        SheetLayoutSnapshot now = snap.Value;
                        if (!now.Equals(old))
                        {
                            Logger.RunWithTaskContext(projectId, taskId, attemptNo, () =>
                            {
                                CompareAndLogLayoutChanges(old, now, ws, suppressLayoutLog: false, trigger: LayoutChangeTrigger.Other);
                            });
                            _layoutSnapshots[key] = now;
                        }
                    }
                    catch (Exception exSheet)
                    {
                        WriteDiagnostic("FlushPendingSelectedSheetsLayoutDiffsForTask sheet error: " + exSheet.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteDiagnostic("FlushPendingSelectedSheetsLayoutDiffsForTask: " + ex.Message);
            }
        }

        private void CompareAndLogLayoutChanges(SheetLayoutSnapshot oldS, SheetLayoutSnapshot newS, Excel.Worksheet ws, bool suppressLayoutLog, LayoutChangeTrigger trigger)
        {
            string sheetName = "";
            try
            {
                sheetName = ws.Name ?? "";
            }
            catch
            {
                sheetName = "?";
            }

            if (suppressLayoutLog)
            {
                WriteDiagnostic($"Layout diff ignored once (trigger={trigger}, sheet={sheetName})");
            }

            if (!suppressLayoutLog && oldS.PrintArea != newS.PrintArea)
                Logger.LogOperation("SetPrintArea", $"{sheetName}!{newS.PrintArea}");

            if (!suppressLayoutLog && (oldS.PrintTitleRows != newS.PrintTitleRows || oldS.PrintTitleColumns != newS.PrintTitleColumns))
                Logger.LogOperation("SetPrintTitle", $"{sheetName}!Rows:{newS.PrintTitleRows};Cols:{newS.PrintTitleColumns}");

            bool headerChanged =
                oldS.LeftHeader != newS.LeftHeader || oldS.CenterHeader != newS.CenterHeader || oldS.RightHeader != newS.RightHeader ||
                oldS.LeftFooter != newS.LeftFooter || oldS.CenterFooter != newS.CenterFooter || oldS.RightFooter != newS.RightFooter;
            if (!suppressLayoutLog && headerChanged)
                Logger.LogOperation("SetHeaderFooter", $"{sheetName}!HF");

            if (!suppressLayoutLog && oldS.Orientation != newS.Orientation)
                Logger.LogOperation("SetPageOrientation", $"{sheetName}!Orientation={newS.Orientation}");

            bool marginChanged =
                !MarginEquals(oldS.LeftMargin, newS.LeftMargin) ||
                !MarginEquals(oldS.RightMargin, newS.RightMargin) ||
                !MarginEquals(oldS.TopMargin, newS.TopMargin) ||
                !MarginEquals(oldS.BottomMargin, newS.BottomMargin);
            if (!suppressLayoutLog && marginChanged)
                Logger.LogOperation("SetPageMargins", $"{sheetName}!L:{newS.LeftMargin};R:{newS.RightMargin};T:{newS.TopMargin};B:{newS.BottomMargin}");

            bool scalingChanged =
                oldS.Zoom != newS.Zoom ||
                oldS.PaperSize != newS.PaperSize ||
                oldS.FitToPagesWide != newS.FitToPagesWide ||
                oldS.FitToPagesTall != newS.FitToPagesTall ||
                oldS.BlackAndWhite != newS.BlackAndWhite ||
                oldS.Draft != newS.Draft;
            if (!suppressLayoutLog && scalingChanged)
                Logger.LogOperation("SetPageScaling", $"{sheetName}!Zoom={newS.Zoom};Paper={newS.PaperSize};FitW={newS.FitToPagesWide};FitT={newS.FitToPagesTall};BW={newS.BlackAndWhite};Draft={newS.Draft}");

            if (!suppressLayoutLog && (oldS.HPageBreakCount != newS.HPageBreakCount || oldS.VPageBreakCount != newS.VPageBreakCount))
                Logger.LogOperation("SetPageBreak", $"{sheetName}!H={newS.HPageBreakCount};V={newS.VPageBreakCount}");

            if (!suppressLayoutLog && oldS.TableStyleSignature != newS.TableStyleSignature)
                Logger.LogOperation("SetTableStyle", $"{sheetName}!{newS.TableStyleSignature}");

            if (!suppressLayoutLog && oldS.TableRangeSignature != newS.TableRangeSignature)
                Logger.LogOperation("ResizeTable", $"{sheetName}!{newS.TableRangeSignature}");

            if (!suppressLayoutLog && oldS.SortFilterSignature != newS.SortFilterSignature)
                Logger.LogOperation("SortOrFilter", $"{sheetName}!{newS.SortFilterSignature}");

            if (!suppressLayoutLog && oldS.HyperlinkSignature != newS.HyperlinkSignature)
                Logger.LogOperation("InsertHyperlink", $"{sheetName}!HyperlinkChanged");

            if (!suppressLayoutLog && oldS.ShapeCount != newS.ShapeCount)
            {
                if (newS.ShapeCount > oldS.ShapeCount)
                    Logger.LogOperation("InsertShapeOrImage", $"{sheetName}!Count:{oldS.ShapeCount}->{newS.ShapeCount}");
                else
                    Logger.LogOperation("DeleteShapeOrImage", $"{sheetName}!Count:{oldS.ShapeCount}->{newS.ShapeCount}");
            }
            else if (!suppressLayoutLog && oldS.ShapeGeometrySignature != newS.ShapeGeometrySignature)
            {
                Logger.LogOperation("MoveOrResizeShape", $"{sheetName}!Count={newS.ShapeCount}");
            }

            if (!suppressLayoutLog && oldS.NamedRangeSignature != newS.NamedRangeSignature)
                Logger.LogOperation("ManageNamedRange", $"{sheetName}!NamedRangeChanged");

            if (!suppressLayoutLog && oldS.ExternalDataSignature != newS.ExternalDataSignature)
                Logger.LogOperation("ImportExternalData", $"{sheetName}!ExternalDataChanged");

            if (!suppressLayoutLog && oldS.ConditionalFormatSignature != newS.ConditionalFormatSignature)
                Logger.LogOperation("AddConditionalFormat", $"{sheetName}!ConditionalFormatChanged");

            if (!suppressLayoutLog && oldS.UsedRowCount != newS.UsedRowCount)
            {
                if (newS.UsedRowCount > oldS.UsedRowCount)
                    Logger.LogOperation("InsertRows", $"{sheetName}!Rows:{oldS.UsedRowCount}->{newS.UsedRowCount}");
                else
                    Logger.LogOperation("DeleteRows", $"{sheetName}!Rows:{oldS.UsedRowCount}->{newS.UsedRowCount}");
            }

            if (!suppressLayoutLog && oldS.UsedColumnCount != newS.UsedColumnCount)
            {
                if (newS.UsedColumnCount > oldS.UsedColumnCount)
                    Logger.LogOperation("InsertColumns", $"{sheetName}!Cols:{oldS.UsedColumnCount}->{newS.UsedColumnCount}");
                else
                    Logger.LogOperation("DeleteColumns", $"{sheetName}!Cols:{oldS.UsedColumnCount}->{newS.UsedColumnCount}");
            }

            if (!suppressLayoutLog && oldS.CellFormatSignature != newS.CellFormatSignature)
                Logger.LogOperation("EditCellFormat", $"{sheetName}!UsedRangeFormatChanged");

            bool freezeOldKnown = oldS.FreezePanes.HasValue;
            bool freezeNewKnown = newS.FreezePanes.HasValue;
            if (freezeOldKnown && freezeNewKnown)
            {
                if (!suppressLayoutLog && (oldS.FreezePanes != newS.FreezePanes ||
                    oldS.SplitRow != newS.SplitRow ||
                    oldS.SplitColumn != newS.SplitColumn))
                {
                    Logger.LogOperation("SetFreezePanes", $"{sheetName}!Freeze={newS.FreezePanes};SplitRow={newS.SplitRow};SplitCol={newS.SplitColumn}");
                }
            }
        }

        private static bool MarginEquals(double a, double b)
        {
            if (double.IsNaN(a) && double.IsNaN(b)) return true;
            if (double.IsNaN(a) || double.IsNaN(b)) return false;
            return Math.Abs(a - b) < 0.0001;
        }

        private struct SheetLayoutSnapshot : IEquatable<SheetLayoutSnapshot>
        {
            public string PrintArea;
            public string PrintTitleRows;
            public string PrintTitleColumns;
            public int Orientation;
            public double LeftMargin, RightMargin, TopMargin, BottomMargin;
            public string LeftHeader, CenterHeader, RightHeader;
            public string LeftFooter, CenterFooter, RightFooter;
            public int Zoom;
            public int PaperSize;
            public int FitToPagesWide;
            public int FitToPagesTall;
            public bool BlackAndWhite;
            public bool Draft;
            public int HPageBreakCount;
            public int VPageBreakCount;
            public string TableStyleSignature;
            public string TableRangeSignature;
            public string SortFilterSignature;
            public int ShapeCount;
            public string ShapeGeometrySignature;
            public string NamedRangeSignature;
            public string ExternalDataSignature;
            public string ConditionalFormatSignature;
            public int UsedRowCount;
            public int UsedColumnCount;
            public string CellFormatSignature;
            public string HyperlinkSignature;
            public bool? FreezePanes;
            public int? SplitRow;
            public int? SplitColumn;

            public bool Equals(SheetLayoutSnapshot other)
            {
                return PrintArea == other.PrintArea
                    && PrintTitleRows == other.PrintTitleRows
                    && PrintTitleColumns == other.PrintTitleColumns
                    && Orientation == other.Orientation
                    && MarginEquals(LeftMargin, other.LeftMargin)
                    && MarginEquals(RightMargin, other.RightMargin)
                    && MarginEquals(TopMargin, other.TopMargin)
                    && MarginEquals(BottomMargin, other.BottomMargin)
                    && LeftHeader == other.LeftHeader
                    && CenterHeader == other.CenterHeader
                    && RightHeader == other.RightHeader
                    && LeftFooter == other.LeftFooter
                    && CenterFooter == other.CenterFooter
                    && RightFooter == other.RightFooter
                    && Zoom == other.Zoom
                    && PaperSize == other.PaperSize
                    && FitToPagesWide == other.FitToPagesWide
                    && FitToPagesTall == other.FitToPagesTall
                    && BlackAndWhite == other.BlackAndWhite
                    && Draft == other.Draft
                    && HPageBreakCount == other.HPageBreakCount
                    && VPageBreakCount == other.VPageBreakCount
                    && TableStyleSignature == other.TableStyleSignature
                    && TableRangeSignature == other.TableRangeSignature
                    && SortFilterSignature == other.SortFilterSignature
                    && ShapeCount == other.ShapeCount
                    && ShapeGeometrySignature == other.ShapeGeometrySignature
                    && NamedRangeSignature == other.NamedRangeSignature
                    && ExternalDataSignature == other.ExternalDataSignature
                    && ConditionalFormatSignature == other.ConditionalFormatSignature
                    && UsedRowCount == other.UsedRowCount
                    && UsedColumnCount == other.UsedColumnCount
                    && CellFormatSignature == other.CellFormatSignature
                    && HyperlinkSignature == other.HyperlinkSignature
                    && FreezePanes == other.FreezePanes
                    && SplitRow == other.SplitRow
                    && SplitColumn == other.SplitColumn;
            }

            public override bool Equals(object obj)
            {
                return obj is SheetLayoutSnapshot other && Equals(other);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    int hc = PrintArea != null ? PrintArea.GetHashCode() : 0;
                    hc = (hc * 397) ^ Orientation;
                    return hc;
                }
            }
        }
    }
}
