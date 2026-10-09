using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        private string _freshBaselineSheetKey;
        private int _freshBaselineStamp;
        private const int FreshBaselineReuseMs = 2000;

        private void RememberFreshBaseline(string sheetKey)
        {
            _freshBaselineSheetKey = sheetKey;
            _freshBaselineStamp = Environment.TickCount;
        }

        private void InvalidateFreshBaseline()
        {
            _freshBaselineSheetKey = null;
        }

        private bool HasFreshBaseline(string sheetKey, out int ageMs)
        {
            ageMs = 0;
            if (string.IsNullOrEmpty(sheetKey) || sheetKey != _freshBaselineSheetKey)
                return false;
            ageMs = unchecked(Environment.TickCount - _freshBaselineStamp);
            return ageMs >= 0 && ageMs < FreshBaselineReuseMs;
        }

        private void InitializeLayoutSnapshotsForAllOpenWorkbooks(string reason)
        {
            if (Application == null) return;
            try
            {
                foreach (Excel.Workbook wb in Application.Workbooks)
                {
                    InitializeLayoutSnapshotsForWorkbook(wb, readFreeze: false, reason);
                }
            }
            catch (Exception ex)
            {
                WriteDiagnostic($"InitializeLayoutSnapshotsForAllOpenWorkbooks reason={reason}: {ex.Message}");
            }
        }

        /// <summary>
        /// 表示中のワークシートだけを基準にする。未表示シートは初回表示時に遅延初期化する。
        /// 同じシートを起動・WorkbookOpen・TaskStart が短時間に連続しても、変更がなければ再取得しない。
        /// </summary>
        private void InitializeLayoutSnapshotsForWorkbook(Excel.Workbook workbook, bool readFreeze, string reason)
        {
            if (workbook == null) return;
            try
            {
                var ws = workbook.ActiveSheet as Excel.Worksheet;
                if (ws == null)
                {
                    WriteDiagnostic($"LayoutSnapshot active skipped reason={reason} (no worksheet)");
                    return;
                }

                string key = GetSheetKey(ws);
                if (HasFreshBaseline(key, out int ageMs) && _layoutSnapshots.ContainsKey(key))
                {
                    WriteDiagnostic(
                        $"LayoutSnapshot skipped reason={reason} sheet={SafeWorksheetName(ws)} freshBaselineAge={ageMs}ms");
                    WriteDiagnostic(WithOpenToken("Baseline completed reason=" + reason + " sheet=" + SafeWorksheetName(ws)));
                    return;
                }

                var sw = Stopwatch.StartNew();
                TryStoreSnapshot(ws, readFreeze, logChanges: false);
                if (_layoutSnapshots.ContainsKey(key))
                    RememberFreshBaseline(key);
                WriteDiagnostic(
                    $"LayoutSnapshot active reason={reason} sheet={SafeWorksheetName(ws)} elapsed={sw.ElapsedMilliseconds}ms");
                if (_layoutSnapshots.ContainsKey(key))
                    WriteDiagnostic(WithOpenToken("Baseline completed reason=" + reason + " sheet=" + SafeWorksheetName(ws)));
            }
            catch (Exception ex)
            {
                WriteDiagnostic($"InitializeLayoutSnapshotsForWorkbook reason={reason}: {ex.Message}");
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
                TryStoreSnapshot(ws, readFreeze: true, logChanges: true, trigger: LayoutChangeTrigger.SheetActivate, reportFirstStore: true);
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
                TryStoreSnapshot(ws, readFreeze: true, logChanges: true, reportFirstStore: true);
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
                TryStoreSnapshot(ws, readFreeze: true, logChanges: false, trigger: LayoutChangeTrigger.WindowActivate, reportFirstStore: true);
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
                InitializeLayoutSnapshotsForWorkbook(wb, readFreeze: false, "NewWorkbook");
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

        private static string SafeWorksheetName(Excel.Worksheet ws)
        {
            try
            {
                return ws?.Name ?? "?";
            }
            catch
            {
                return "?";
            }
        }

        private void TryStoreSnapshot(Excel.Worksheet ws, bool readFreeze, bool logChanges, LayoutChangeTrigger trigger = LayoutChangeTrigger.Other, bool reportFirstStore = false)
        {
            var sw = Stopwatch.StartNew();
            SheetLayoutSnapshot? snap = BuildSnapshot(ws, readFreeze);
            if (snap == null) return;

            string key = GetSheetKey(ws);

            if (!_layoutSnapshots.TryGetValue(key, out SheetLayoutSnapshot old))
            {
                _layoutSnapshots[key] = snap.Value;
                if (reportFirstStore)
                {
                    WriteDiagnostic(
                        $"LayoutSnapshot deferred sheet={SafeWorksheetName(ws)} elapsed={sw.ElapsedMilliseconds}ms trigger={trigger}");
                }
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

        private sealed class SnapshotReadCache
        {
            public readonly Dictionary<string, List<string>> WorkbookNameParts =
                new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 未取得の項目はビットを立てない。初期値と実値を差分ログにしないため。
        /// PageSetup と改ページは、非印刷タスクでも無断変更が違反になるので起動基準に含める。
        /// </summary>
        [Flags]
        private enum SnapshotFields
        {
            None = 0,
            PageSetup = 1,
            PageBreaks = 2,
            Tables = 4,
            SortFilter = 8,
            Shapes = 16,
            NamedRanges = 32,
            ExternalData = 64,
            ConditionalFormats = 128,
            UsedRange = 256,
            CellFormat = 512,
            Hyperlinks = 1024,
            Freeze = 2048
        }

        private const SnapshotFields TrackedSnapshotFields =
            SnapshotFields.PageSetup
            | SnapshotFields.PageBreaks
            | SnapshotFields.Tables
            | SnapshotFields.SortFilter
            | SnapshotFields.Shapes
            | SnapshotFields.NamedRanges
            | SnapshotFields.ExternalData
            | SnapshotFields.ConditionalFormats
            | SnapshotFields.UsedRange
            | SnapshotFields.CellFormat
            | SnapshotFields.Hyperlinks;

        private static long MarkTiming(Stopwatch sw, long mark, string name, List<string> steps)
        {
            long now = sw.ElapsedMilliseconds;
            steps.Add(name + "=" + (now - mark).ToString(CultureInfo.InvariantCulture));
            return now;
        }

        private static SheetLayoutSnapshot? BuildSnapshot(Excel.Worksheet ws, bool readFreeze, SnapshotReadCache cache = null)
        {
            var steps = new List<string>(12);
            var sw = Stopwatch.StartNew();
            long mark = 0;
            try
            {
                var s = new SheetLayoutSnapshot();
                s.Present = TrackedSnapshotFields;

                Excel.PageSetup ps = ws.PageSetup;
                s.PrintArea = SafeGet(() => ps.PrintArea);
                s.PrintTitleRows = SafeGet(() => ps.PrintTitleRows);
                s.PrintTitleColumns = SafeGet(() => ps.PrintTitleColumns);
                s.Orientation = SafeGetInt(() => (int)ps.Orientation);
                s.LeftMargin = SafeGetDouble(() => ps.LeftMargin);
                s.RightMargin = SafeGetDouble(() => ps.RightMargin);
                s.TopMargin = SafeGetDouble(() => ps.TopMargin);
                s.BottomMargin = SafeGetDouble(() => ps.BottomMargin);
                s.LeftHeader = SafeGet(() => ps.LeftHeader);
                s.CenterHeader = SafeGet(() => ps.CenterHeader);
                s.RightHeader = SafeGet(() => ps.RightHeader);
                s.LeftFooter = SafeGet(() => ps.LeftFooter);
                s.CenterFooter = SafeGet(() => ps.CenterFooter);
                s.RightFooter = SafeGet(() => ps.RightFooter);
                s.Zoom = SafeGetZoom(ps);
                s.PaperSize = SafeGetInt(() => (int)ps.PaperSize);
                s.FitToPagesWide = SafeGetInt(() => ps.FitToPagesWide);
                s.FitToPagesTall = SafeGetInt(() => ps.FitToPagesTall);
                s.BlackAndWhite = SafeGetBool(() => ps.BlackAndWhite);
                s.Draft = SafeGetBool(() => ps.Draft);
                mark = MarkTiming(sw, mark, "PageSetup", steps);

                s.HPageBreakCount = SafeGetHBreakCount(ws);
                s.VPageBreakCount = SafeGetVBreakCount(ws);
                mark = MarkTiming(sw, mark, "PageBreaks", steps);

                s.TableStyleSignature = BuildTableStyleSignature(ws);
                s.TableRangeSignature = BuildTableRangeSignature(ws);
                mark = MarkTiming(sw, mark, "Tables", steps);

                s.SortFilterSignature = BuildSortFilterSignature(ws);
                mark = MarkTiming(sw, mark, "SortFilter", steps);

                s.ShapeCount = SafeGetShapeCount(ws);
                s.ShapeGeometrySignature = BuildShapeGeometrySignature(ws);
                mark = MarkTiming(sw, mark, "Shapes", steps);

                s.NamedRangeSignature = BuildNamedRangeSignature(ws, cache);
                mark = MarkTiming(sw, mark, "NamedRanges", steps);

                s.ExternalDataSignature = BuildExternalDataSignature(ws);
                mark = MarkTiming(sw, mark, "ExternalData", steps);

                Excel.Range used = null;
                try { used = ws.UsedRange; } catch { }
                s.UsedRowCount = SafeGetUsedRangeRowCount(used);
                s.UsedColumnCount = SafeGetUsedRangeColumnCount(used);
                mark = MarkTiming(sw, mark, "UsedRange", steps);

                s.CellFormatSignature = BuildCellFormatSignature(used);
                mark = MarkTiming(sw, mark, "CellFormat", steps);

                s.ConditionalFormatSignature = BuildConditionalFormatSignature(used);
                mark = MarkTiming(sw, mark, "ConditionalFormats", steps);

                s.HyperlinkSignature = BuildHyperlinkSignature(ws);
                mark = MarkTiming(sw, mark, "Hyperlinks", steps);

                if (readFreeze && TryGetFreezeForActiveSheet(ws, out bool freeze, out int splitRow, out int splitCol))
                {
                    s.Present |= SnapshotFields.Freeze;
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
                MarkTiming(sw, mark, "Freeze", steps);

                if (sw.ElapsedMilliseconds >= 30)
                {
                    WriteDiagnostic(
                        "BuildSnapshot sheet=" + SafeWorksheetName(ws)
                        + " total=" + sw.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                        + "ms " + string.Join(" ", steps));
                }

                return s;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[LayoutMonitoring] BuildSnapshot: " + ex.Message);
                if (sw.ElapsedMilliseconds >= 30)
                {
                    WriteDiagnostic(
                        "BuildSnapshot failed sheet=" + SafeWorksheetName(ws)
                        + " total=" + sw.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                        + "ms " + string.Join(" ", steps)
                        + " error=" + ex.Message);
                }
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

        private static string BuildNamedRangeSignature(Excel.Worksheet ws, SnapshotReadCache cache)
        {
            var parts = new List<string>();
            try
            {
                var wb = ws.Parent as Excel.Workbook;
                if (wb != null)
                {
                    string wbKey = GetWorkbookKey(wb);
                    List<string> workbookParts = null;
                    if (cache == null || !cache.WorkbookNameParts.TryGetValue(wbKey, out workbookParts))
                    {
                        workbookParts = new List<string>();
                        foreach (Excel.Name n in wb.Names)
                        {
                            try
                            {
                                string name = SafeGet(() => n.Name);
                                string refersTo = SafeGet(() => n.RefersTo);
                                workbookParts.Add("WB:" + name + ":" + refersTo);
                            }
                            catch { }
                        }

                        if (cache != null)
                            cache.WorkbookNameParts[wbKey] = workbookParts;
                    }

                    if (workbookParts != null)
                        parts.AddRange(workbookParts);
                }

                foreach (Excel.Name n in ws.Names)
                {
                    try
                    {
                        string name = SafeGet(() => n.Name);
                        string refersTo = SafeGet(() => n.RefersTo);
                        parts.Add("WS:" + name + ":" + refersTo);
                    }
                    catch { }
                }
            }
            catch { }

            parts.Sort(StringComparer.Ordinal);
            return string.Join("|", parts);
        }

        /// <summary>
        /// 名前定義シグネチャから定義名の集合だけを取り出し、追加・削除があるか判定する。
        /// RefersTo だけの変化（SORT/スピル等の自動更新）は破壊的操作にしない。
        /// 署名要素形式: "WB:名前:参照先" / "WS:名前:参照先"
        /// </summary>
        private static bool HasNamedRangeNameSetChanged(string oldSignature, string newSignature)
        {
            var oldNames = ExtractNamedRangeNameKeys(oldSignature);
            var newNames = ExtractNamedRangeNameKeys(newSignature);
            if (oldNames.Count != newNames.Count)
                return true;
            foreach (string name in oldNames)
            {
                if (!newNames.Contains(name))
                    return true;
            }
            return false;
        }

        private static HashSet<string> ExtractNamedRangeNameKeys(string signature)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(signature))
                return keys;

            foreach (string part in signature.Split('|'))
            {
                if (string.IsNullOrEmpty(part))
                    continue;

                // "WB:氏名:=Sheet!$A$1" → scope=WB, name=氏名
                int first = part.IndexOf(':');
                if (first < 0 || first >= part.Length - 1)
                    continue;
                int second = part.IndexOf(':', first + 1);
                string scope = part.Substring(0, first);
                string name = second > first
                    ? part.Substring(first + 1, second - first - 1)
                    : part.Substring(first + 1);
                if (string.IsNullOrEmpty(name))
                    continue;
                keys.Add(scope + ":" + name);
            }
            return keys;
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

        private static string BuildConditionalFormatSignature(Excel.Range used)
        {
            var parts = new List<string>();
            try
            {
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

        private static int SafeGetUsedRangeRowCount(Excel.Range used)
        {
            try
            {
                if (used == null) return 0;
                return used.Rows?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static int SafeGetUsedRangeColumnCount(Excel.Range used)
        {
            try
            {
                if (used == null) return 0;
                return used.Columns?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static string BuildCellFormatSignature(Excel.Range used)
        {
            try
            {
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
            var capturedSheetKeys = new HashSet<string>(StringComparer.Ordinal);
            var cache = new SnapshotReadCache();
            FlushPendingSelectedSheetsLayoutDiffsForTask(projectId, taskId, attemptNo, capturedSheetKeys, cache);
            FlushPendingWorkbookSheetsBoundaryDiffsSinglePass(projectId, taskId, attemptNo, capturedSheetKeys, cache);
        }

        /// <summary>
        /// 全ブック各シートについて、旧スナップショットと現在状態を1回の <c>BuildSnapshot</c> で比較し、
        /// 差分カテゴリをまとめてログしてからスナップショットを更新する。
        /// </summary>
        private void FlushPendingWorkbookSheetsBoundaryDiffsSinglePass(
            int projectId,
            int taskId,
            int attemptNo,
            HashSet<string> alreadyCapturedSheetKeys,
            SnapshotReadCache cache)
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
                                if (alreadyCapturedSheetKeys != null && alreadyCapturedSheetKeys.Contains(key))
                                    continue;

                                SheetLayoutSnapshot? snap = BuildSnapshot(ws, readFreeze: false, cache);
                                if (snap == null) continue;

                                if (!_layoutSnapshots.TryGetValue(key, out SheetLayoutSnapshot old))
                                {
                                    _layoutSnapshots[key] = snap.Value;
                                    continue;
                                }

                                SheetLayoutSnapshot now = snap.Value;
                                bool rowChanged = CanCompare(old, now, SnapshotFields.UsedRange) && old.UsedRowCount != now.UsedRowCount;
                                bool colChanged = CanCompare(old, now, SnapshotFields.UsedRange) && old.UsedColumnCount != now.UsedColumnCount;
                                bool sortFilterChanged = CanCompare(old, now, SnapshotFields.SortFilter) && old.SortFilterSignature != now.SortFilterSignature;
                                bool tableStyleChanged = CanCompare(old, now, SnapshotFields.Tables) && old.TableStyleSignature != now.TableStyleSignature;
                                bool shapeCountChanged = CanCompare(old, now, SnapshotFields.Shapes) && old.ShapeCount != now.ShapeCount;
                                bool shapeGeomChanged = CanCompare(old, now, SnapshotFields.Shapes) && old.ShapeGeometrySignature != now.ShapeGeometrySignature;
                                bool hyperlinkChanged = CanCompare(old, now, SnapshotFields.Hyperlinks) && old.HyperlinkSignature != now.HyperlinkSignature;

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
                                    if (!IsLikelyFormatOnlyUsedRangeDrift(old, now))
                                    {
                                        if (now.UsedRowCount > old.UsedRowCount)
                                            Logger.LogOperation("InsertRows", $"{sheetName}!Rows:{old.UsedRowCount}->{now.UsedRowCount};Trigger=BoundaryFlush");
                                        else
                                            Logger.LogOperation("DeleteRows", $"{sheetName}!Rows:{old.UsedRowCount}->{now.UsedRowCount};Trigger=BoundaryFlush");
                                    }
                                }

                                if (colChanged)
                                {
                                    if (!IsLikelyFormatOnlyUsedRangeDrift(old, now))
                                    {
                                        if (now.UsedColumnCount > old.UsedColumnCount)
                                            Logger.LogOperation("InsertColumns", $"{sheetName}!Cols:{old.UsedColumnCount}->{now.UsedColumnCount};Trigger=BoundaryFlush");
                                        else
                                            Logger.LogOperation("DeleteColumns", $"{sheetName}!Cols:{old.UsedColumnCount}->{now.UsedColumnCount};Trigger=BoundaryFlush");
                                    }
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
        private void FlushPendingSelectedSheetsLayoutDiffsForTask(
            int projectId,
            int taskId,
            int attemptNo,
            HashSet<string> capturedSheetKeys,
            SnapshotReadCache cache)
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
                        SheetLayoutSnapshot? snap = BuildSnapshot(ws, readFreeze: true, cache);
                        if (snap == null) continue;

                        if (!_layoutSnapshots.TryGetValue(key, out SheetLayoutSnapshot old))
                        {
                            _layoutSnapshots[key] = snap.Value;
                            if (capturedSheetKeys != null)
                                capturedSheetKeys.Add(key);
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
                        if (capturedSheetKeys != null)
                            capturedSheetKeys.Add(key);
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

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.PageSetup) && oldS.PrintArea != newS.PrintArea)
                Logger.LogOperation("SetPrintArea", $"{sheetName}!{newS.PrintArea}");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.PageSetup) && (oldS.PrintTitleRows != newS.PrintTitleRows || oldS.PrintTitleColumns != newS.PrintTitleColumns))
                Logger.LogOperation("SetPrintTitle", $"{sheetName}!Rows:{newS.PrintTitleRows};Cols:{newS.PrintTitleColumns}");

            bool headerChanged =
                oldS.LeftHeader != newS.LeftHeader || oldS.CenterHeader != newS.CenterHeader || oldS.RightHeader != newS.RightHeader ||
                oldS.LeftFooter != newS.LeftFooter || oldS.CenterFooter != newS.CenterFooter || oldS.RightFooter != newS.RightFooter;
            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.PageSetup) && headerChanged)
                Logger.LogOperation("SetHeaderFooter", $"{sheetName}!HF");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.PageSetup) && oldS.Orientation != newS.Orientation)
                Logger.LogOperation("SetPageOrientation", $"{sheetName}!Orientation={newS.Orientation}");

            bool marginChanged =
                !MarginEquals(oldS.LeftMargin, newS.LeftMargin) ||
                !MarginEquals(oldS.RightMargin, newS.RightMargin) ||
                !MarginEquals(oldS.TopMargin, newS.TopMargin) ||
                !MarginEquals(oldS.BottomMargin, newS.BottomMargin);
            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.PageSetup) && marginChanged)
                Logger.LogOperation("SetPageMargins", $"{sheetName}!L:{newS.LeftMargin};R:{newS.RightMargin};T:{newS.TopMargin};B:{newS.BottomMargin}");

            bool scalingChanged =
                oldS.Zoom != newS.Zoom ||
                oldS.PaperSize != newS.PaperSize ||
                oldS.FitToPagesWide != newS.FitToPagesWide ||
                oldS.FitToPagesTall != newS.FitToPagesTall ||
                oldS.BlackAndWhite != newS.BlackAndWhite ||
                oldS.Draft != newS.Draft;
            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.PageSetup) && scalingChanged)
                Logger.LogOperation("SetPageScaling", $"{sheetName}!Zoom={newS.Zoom};Paper={newS.PaperSize};FitW={newS.FitToPagesWide};FitT={newS.FitToPagesTall};BW={newS.BlackAndWhite};Draft={newS.Draft}");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.PageBreaks) && (oldS.HPageBreakCount != newS.HPageBreakCount || oldS.VPageBreakCount != newS.VPageBreakCount))
                Logger.LogOperation("SetPageBreak", $"{sheetName}!H={newS.HPageBreakCount};V={newS.VPageBreakCount}");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.Tables) && oldS.TableStyleSignature != newS.TableStyleSignature)
                Logger.LogOperation("SetTableStyle", $"{sheetName}!{newS.TableStyleSignature}");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.Tables) && oldS.TableRangeSignature != newS.TableRangeSignature)
                Logger.LogOperation("ResizeTable", $"{sheetName}!{newS.TableRangeSignature}");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.SortFilter) && oldS.SortFilterSignature != newS.SortFilterSignature)
                Logger.LogOperation("SortOrFilter", $"{sheetName}!{newS.SortFilterSignature}");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.Hyperlinks) && oldS.HyperlinkSignature != newS.HyperlinkSignature)
                Logger.LogOperation("InsertHyperlink", $"{sheetName}!HyperlinkChanged");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.Shapes) && oldS.ShapeCount != newS.ShapeCount)
            {
                if (newS.ShapeCount > oldS.ShapeCount)
                    Logger.LogOperation("InsertShapeOrImage", $"{sheetName}!Count:{oldS.ShapeCount}->{newS.ShapeCount}");
                else
                    Logger.LogOperation("DeleteShapeOrImage", $"{sheetName}!Count:{oldS.ShapeCount}->{newS.ShapeCount}");
            }
            else if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.Shapes) && oldS.ShapeGeometrySignature != newS.ShapeGeometrySignature)
            {
                Logger.LogOperation("MoveOrResizeShape", $"{sheetName}!Count={newS.ShapeCount}");
            }

            // 参照先(RefersTo)だけの自動更新は無視し、定義名の追加・削除があるときだけ記録する
            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.NamedRanges)
                && HasNamedRangeNameSetChanged(oldS.NamedRangeSignature, newS.NamedRangeSignature))
                Logger.LogOperation("ManageNamedRange", $"{sheetName}!NamedRangeChanged");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.ExternalData) && oldS.ExternalDataSignature != newS.ExternalDataSignature)
                Logger.LogOperation("ImportExternalData", $"{sheetName}!ExternalDataChanged");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.ConditionalFormats) && oldS.ConditionalFormatSignature != newS.ConditionalFormatSignature)
                Logger.LogOperation("AddConditionalFormat", $"{sheetName}!ConditionalFormatChanged");

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.UsedRange) && oldS.UsedRowCount != newS.UsedRowCount)
            {
                if (!IsLikelyFormatOnlyUsedRangeDrift(oldS, newS))
                {
                    if (newS.UsedRowCount > oldS.UsedRowCount)
                        Logger.LogOperation("InsertRows", $"{sheetName}!Rows:{oldS.UsedRowCount}->{newS.UsedRowCount}");
                    else
                        Logger.LogOperation("DeleteRows", $"{sheetName}!Rows:{oldS.UsedRowCount}->{newS.UsedRowCount}");
                }
            }

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.UsedRange) && oldS.UsedColumnCount != newS.UsedColumnCount)
            {
                if (!IsLikelyFormatOnlyUsedRangeDrift(oldS, newS))
                {
                    if (newS.UsedColumnCount > oldS.UsedColumnCount)
                        Logger.LogOperation("InsertColumns", $"{sheetName}!Cols:{oldS.UsedColumnCount}->{newS.UsedColumnCount}");
                    else
                        Logger.LogOperation("DeleteColumns", $"{sheetName}!Cols:{oldS.UsedColumnCount}->{newS.UsedColumnCount}");
                }
            }

            if (!suppressLayoutLog && CanCompare(oldS, newS, SnapshotFields.CellFormat) && oldS.CellFormatSignature != newS.CellFormatSignature)
                Logger.LogOperation("EditCellFormat", $"{sheetName}!UsedRangeFormatChanged");

            bool freezeOldKnown = CanCompare(oldS, newS, SnapshotFields.Freeze) && oldS.FreezePanes.HasValue;
            bool freezeNewKnown = CanCompare(oldS, newS, SnapshotFields.Freeze) && newS.FreezePanes.HasValue;
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

        private static bool CanCompare(SheetLayoutSnapshot oldS, SheetLayoutSnapshot newS, SnapshotFields field)
        {
            return (oldS.Present & field) == field && (newS.Present & field) == field;
        }

        private static bool MarginEquals(double a, double b)
        {
            if (double.IsNaN(a) && double.IsNaN(b)) return true;
            if (double.IsNaN(a) || double.IsNaN(b)) return false;
            return Math.Abs(a - b) < 0.0001;
        }

        /// <summary>
        /// セル書式変更に伴い UsedRange の行数/列数だけがわずかに変わった誤検知を抑える。
        /// </summary>
        private static bool IsLikelyFormatOnlyUsedRangeDrift(SheetLayoutSnapshot oldS, SheetLayoutSnapshot newS)
        {
            if (!CanCompare(oldS, newS, SnapshotFields.CellFormat) || !CanCompare(oldS, newS, SnapshotFields.UsedRange))
                return false;
            if (oldS.CellFormatSignature == newS.CellFormatSignature)
                return false;

            if (Math.Abs(oldS.UsedRowCount - newS.UsedRowCount) > 1)
                return false;
            if (Math.Abs(oldS.UsedColumnCount - newS.UsedColumnCount) > 1)
                return false;

            if (oldS.TableStyleSignature != newS.TableStyleSignature) return false;
            if (oldS.TableRangeSignature != newS.TableRangeSignature) return false;
            if (oldS.SortFilterSignature != newS.SortFilterSignature) return false;
            if (oldS.ShapeCount != newS.ShapeCount) return false;
            if (oldS.ShapeGeometrySignature != newS.ShapeGeometrySignature) return false;
            if (oldS.NamedRangeSignature != newS.NamedRangeSignature) return false;
            if (oldS.ExternalDataSignature != newS.ExternalDataSignature) return false;
            if (oldS.ConditionalFormatSignature != newS.ConditionalFormatSignature) return false;
            if (oldS.HyperlinkSignature != newS.HyperlinkSignature) return false;
            if (oldS.HPageBreakCount != newS.HPageBreakCount) return false;
            if (oldS.VPageBreakCount != newS.VPageBreakCount) return false;

            return true;
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
            public SnapshotFields Present;
            public bool? FreezePanes;
            public int? SplitRow;
            public int? SplitColumn;

            public bool Equals(SheetLayoutSnapshot other)
            {
                if (Present != other.Present)
                    return false;
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
