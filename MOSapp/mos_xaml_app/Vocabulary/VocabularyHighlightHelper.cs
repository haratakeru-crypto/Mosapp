using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelListObject = Microsoft.Office.Interop.Excel.ListObject;
using ExcelWindow = Microsoft.Office.Interop.Excel.Window;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>
    /// コーチマーク用の画面座標。単語帳テーブルは B8:F14（指示文には出さない）。
    /// </summary>
    public static class VocabularyHighlightHelper
    {
        public const string VocabTableAddress = "B8:F14";
        const string NativeHighlightShapeName = "MOS_VocabTableHighlight";

        // Excel RGB は BGR
        const int OrangeRgb = 0x20B0FF;

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern uint GetDpiForWindow(IntPtr hWnd);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        /// <summary>
        /// Excel が取れるまで短時間リトライし、物理ピクセル座標のハイライト矩形を返す。
        /// </summary>
        public static IReadOnlyList<Rect> ResolveHighlights(
            IntPtr excelHwnd,
            Func<ExcelApp> getExcelApp,
            string highlightHint)
        {
            ExcelApp excel = null;
            int resolvedTries = 0;
            for (int attempt = 0; attempt < 15; attempt++)
            {
                try { excel = getExcelApp?.Invoke(); } catch { excel = null; }
                if (excel == null)
                {
                    try { excel = (ExcelApp)Marshal.GetActiveObject("Excel.Application"); }
                    catch { excel = null; }
                }

                if (excel != null)
                {
                    IntPtr hwnd = excelHwnd;
                    try
                    {
                        int h = excel.Hwnd;
                        if (h != 0) hwnd = new IntPtr(h);
                    }
                    catch { }

                    var result = ResolveWithExcel(hwnd, excel, highlightHint);
                    if (result.Count > 0) return result;
                    // Excel は取れているのに位置が無いときは待っても変わらない。呼び出し側の再試行に任せる。
                    if (++resolvedTries >= 2) return result;
                }

                Thread.Sleep(100);
            }

            return ResolveWithoutExcel(excelHwnd, highlightHint);
        }

        static IReadOnlyList<Rect> ResolveWithExcel(IntPtr excelHwnd, ExcelApp excelApp, string highlightHint)
        {
            var list = new List<Rect>();
            string hint = highlightHint ?? "";

            if (excelHwnd == IntPtr.Zero || !GetWindowRect(excelHwnd, out RECT wr))
                return list;

            double left = wr.Left;
            double top = wr.Top;
            double width = Math.Max(100, wr.Right - wr.Left);
            double height = Math.Max(100, wr.Bottom - wr.Top);

            if (hint.Equals("Table", StringComparison.OrdinalIgnoreCase))
            {
                Rect? table = TryGetTableScreenRect(excelApp, left, top, width, height);
                if (table.HasValue) list.Add(table.Value);
                return list;
            }

            if (hint.Equals("TableDesignTab", StringComparison.OrdinalIgnoreCase))
            {
                Rect? tab = VocabularyRibbonTabProbe.TryGetTableDesignTabScreenRect(excelHwnd);
                if (tab.HasValue) list.Add(tab.Value);
                return list;
            }

            if (hint.Equals("TableThenDesignTab", StringComparison.OrdinalIgnoreCase))
            {
                Rect? table = TryGetTableScreenRect(excelApp, left, top, width, height);
                if (table.HasValue) list.Add(table.Value);
                Rect? tab = VocabularyRibbonTabProbe.TryGetTableDesignTabScreenRect(excelHwnd);
                if (tab.HasValue) list.Add(tab.Value);
                return list;
            }

            if (hint.Equals("Chart", StringComparison.OrdinalIgnoreCase))
            {
                list.Add(new Rect(left + width * 0.48, top + height * 0.30, width * 0.38, height * 0.34));
                return list;
            }

            if (hint.Equals("ChartDesignTab", StringComparison.OrdinalIgnoreCase))
            {
                Rect? tab = VocabularyRibbonTabProbe.TryGetChartDesignTabScreenRect(excelHwnd);
                if (tab.HasValue) list.Add(tab.Value);
                return list;
            }

            if (hint.Equals("ChartThenDesignTab", StringComparison.OrdinalIgnoreCase))
            {
                list.Add(new Rect(left + width * 0.48, top + height * 0.30, width * 0.38, height * 0.34));
                Rect? tab = VocabularyRibbonTabProbe.TryGetChartDesignTabScreenRect(excelHwnd);
                if (tab.HasValue) list.Add(tab.Value);
                return list;
            }

            if (hint.Equals("FormulaBar", StringComparison.OrdinalIgnoreCase))
            {
                list.Add(new Rect(left + 80, top + 98, Math.Min(520, width * 0.55), 28));
                return list;
            }

            if (hint.StartsWith("File.", StringComparison.OrdinalIgnoreCase))
            {
                list.Add(new Rect(left + 6, top + 26, 64, 34));
                return list;
            }

            // タブ位置は UIA 実座標のみ。割合の仮矩形は使わない。
            return list;
        }

        static IReadOnlyList<Rect> ResolveWithoutExcel(IntPtr excelHwnd, string highlightHint)
        {
            var list = new List<Rect>();
            string hint = highlightHint ?? "";
            if (excelHwnd == IntPtr.Zero || !GetWindowRect(excelHwnd, out RECT wr))
                return list;

            double left = wr.Left;
            double top = wr.Top;
            double width = Math.Max(100, wr.Right - wr.Left);

            if (hint.Equals("Table", StringComparison.OrdinalIgnoreCase))
                return list;

            if (hint.Equals("TableDesignTab", StringComparison.OrdinalIgnoreCase)
                || hint.Equals("TableThenDesignTab", StringComparison.OrdinalIgnoreCase))
            {
                Rect? tab = VocabularyRibbonTabProbe.TryGetTableDesignTabScreenRect(excelHwnd);
                if (tab.HasValue) list.Add(tab.Value);
                return list;
            }

            if (hint.Equals("ChartDesignTab", StringComparison.OrdinalIgnoreCase)
                || hint.Equals("ChartThenDesignTab", StringComparison.OrdinalIgnoreCase))
            {
                Rect? tab = VocabularyRibbonTabProbe.TryGetChartDesignTabScreenRect(excelHwnd);
                if (tab.HasValue) list.Add(tab.Value);
                return list;
            }

            if (hint.Equals("FormulaBar", StringComparison.OrdinalIgnoreCase))
            {
                list.Add(new Rect(left + 80, top + 98, Math.Min(520, width * 0.55), 28));
                return list;
            }

            return list;
        }

        public static ExcelRange TryGetVocabTableRange(ExcelApp excelApp)
        {
            if (excelApp == null) return null;
            try
            {
                ExcelWorksheet ws = excelApp.ActiveSheet as ExcelWorksheet;
                if (ws == null) return null;

                try
                {
                    if (ws.ListObjects != null && ws.ListObjects.Count >= 1)
                        return ws.ListObjects[1].Range;
                }
                catch { }

                return ws.Range[VocabTableAddress];
            }
            catch
            {
                return null;
            }
        }

        static Rect? TryGetTableScreenRect(
            ExcelApp excelApp, double winLeft, double winTop, double winW, double winH)
        {
            ExcelRange range = TryGetVocabTableRange(excelApp);
            if (range == null) return null;

            Rect? r = TryGetRangeScreenRect(excelApp, range, winLeft, winTop, winW, winH, scrollIntoView: false);
            if (r.HasValue) return r;

            r = TryGetRangeScreenRect(excelApp, range, winLeft, winTop, winW, winH, scrollIntoView: true);
            if (r.HasValue) return r;

            // VisibleRange 基準の相対位置（最終手段・シート領域内）
            return TryGetRangeScreenRectViaVisibleRange(excelApp, range, winLeft, winTop, winW, winH);
        }

        static Rect? TryGetRangeScreenRect(
            ExcelApp excelApp,
            ExcelRange range,
            double winLeft,
            double winTop,
            double winW,
            double winH,
            bool scrollIntoView)
        {
            if (excelApp == null || range == null) return null;
            try
            {
                ExcelWindow win = excelApp.ActiveWindow;
                if (win == null) return null;

                if (scrollIntoView)
                {
                    try { win.ScrollRow = Math.Max(1, range.Row - 1); } catch { }
                    try { win.ScrollColumn = Math.Max(1, range.Column - 1); } catch { }
                }

                // 四隅セルで取得（Range.Left/Width より安定）
                ExcelRange ul = (ExcelRange)range.Cells[1, 1];
                ExcelRange lr = (ExcelRange)range.Cells[range.Rows.Count, range.Columns.Count];

                double leftPt = Convert.ToDouble(ul.Left);
                double topPt = Convert.ToDouble(ul.Top);
                double rightPt = Convert.ToDouble(lr.Left) + Convert.ToDouble(lr.Width);
                double bottomPt = Convert.ToDouble(lr.Top) + Convert.ToDouble(lr.Height);

                int x1 = win.PointsToScreenPixelsX((int)Math.Round(leftPt));
                int y1 = win.PointsToScreenPixelsY((int)Math.Round(topPt));
                int x2 = win.PointsToScreenPixelsX((int)Math.Round(rightPt));
                int y2 = win.PointsToScreenPixelsY((int)Math.Round(bottomPt));

                var unscaled = new Rect(
                    Math.Min(x1, x2),
                    Math.Min(y1, y2),
                    Math.Max(8, Math.Abs(x2 - x1)),
                    Math.Max(8, Math.Abs(y2 - y1)));
                if (unscaled.Width >= 16 && unscaled.Height >= 16)
                {
                    // Excel と同じ画面座標で、対象セルの外側まで走査して範囲を合わせる。
                    Rect? snapped = SnapRectToRange(win, range, unscaled);
                    if (snapped.HasValue) return snapped;
                }

                // Excel が論理ピクセルを返す環境向け: 物理ウィンドウ幅と期待サイズで補正
                double dpiScale = 1.0;
                try
                {
                    int hwnd = excelApp.Hwnd;
                    if (hwnd != 0)
                    {
                        uint dpi = GetDpiForWindow(new IntPtr(hwnd));
                        if (dpi > 0) dpiScale = dpi / 96.0;
                    }
                }
                catch { }

                double zoom = 100.0;
                try { zoom = Convert.ToDouble(win.Zoom); } catch { }
                double measuredW = Math.Abs(x2 - x1);
                double expectedPhysW = Convert.ToDouble(range.Width) * (dpiScale * 96.0 / 72.0) * (zoom / 100.0);

                if (dpiScale > 1.05 && measuredW > 8 && expectedPhysW > 8)
                {
                    // 実測が期待物理幅の ~1/dpi なら、論理座標 → 物理へ拡大
                    double ratio = measuredW / expectedPhysW;
                    if (ratio < 0.85 && ratio > 0.4)
                    {
                        // ウィンドウ原点からの相対を拡大して物理スクリーンへ
                        double rx1 = winLeft + (x1 - winLeft) * dpiScale;
                        double ry1 = winTop + (y1 - winTop) * dpiScale;
                        double rx2 = winLeft + (x2 - winLeft) * dpiScale;
                        double ry2 = winTop + (y2 - winTop) * dpiScale;
                        x1 = (int)Math.Round(rx1);
                        y1 = (int)Math.Round(ry1);
                        x2 = (int)Math.Round(rx2);
                        y2 = (int)Math.Round(ry2);
                    }
                }

                var raw = new Rect(
                    Math.Min(x1, x2),
                    Math.Min(y1, y2),
                    Math.Max(8, Math.Abs(x2 - x1)),
                    Math.Max(8, Math.Abs(y2 - y1)));

                if (raw.Width < 16 || raw.Height < 16) return null;

                var clipped = raw;
                clipped.Intersect(new Rect(winLeft, winTop, winW, winH));
                if (clipped.Width >= 16 && clipped.Height >= 16)
                    return clipped;

                double cx = raw.X + raw.Width / 2;
                double cy = raw.Y + raw.Height / 2;
                if (cx >= winLeft && cx <= winLeft + winW && cy >= winTop && cy <= winTop + winH)
                    return raw;

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>RangeFromPoint で穴の四辺を対象セル範囲に合わせる。</summary>
        static Rect? SnapRectToRange(ExcelWindow win, ExcelRange range, Rect approx)
        {
            if (win == null || range == null) return null;
            int row1, col1, row2, col2;
            try
            {
                row1 = range.Row;
                col1 = range.Column;
                row2 = row1 + range.Rows.Count - 1;
                col2 = col1 + range.Columns.Count - 1;
            }
            catch { return null; }

            if (!TryFindInside(win, approx, row1, col1, row2, col2, out int cx, out int cy))
                return null;

            int limitTop = (int)approx.Y - 160;
            int limitBottom = (int)approx.Bottom + 160;
            int limitLeft = (int)approx.X - 160;
            int limitRight = (int)approx.Right + 160;

            int top = FindBoundary(y => IsInRange(win, cx, y, row1, col1, row2, col2), cy, -1, limitTop);
            int bottom = FindBoundary(y => IsInRange(win, cx, y, row1, col1, row2, col2), cy, 1, limitBottom);
            int left = FindBoundary(x => IsInRange(win, x, cy, row1, col1, row2, col2), cx, -1, limitLeft);
            int right = FindBoundary(x => IsInRange(win, x, cy, row1, col1, row2, col2), cx, 1, limitRight);

            int w = right - left;
            int h = bottom - top;
            if (w < 16 || h < 16) return null;
            return new Rect(left, top, w, h);
        }

        static bool TryFindInside(
            ExcelWindow win, Rect approx, int row1, int col1, int row2, int col2, out int x, out int y)
        {
            x = (int)Math.Round(approx.X + approx.Width / 2.0);
            y = (int)Math.Round(approx.Y + approx.Height / 2.0);
            if (IsInRange(win, x, y, row1, col1, row2, col2)) return true;

            int y0 = (int)approx.Y;
            int y1 = (int)approx.Bottom;
            int x0 = (int)approx.X;
            int x1 = (int)approx.Right;
            for (int yy = y0; yy <= y1; yy += 14)
            {
                for (int xx = x0; xx <= x1; xx += 18)
                {
                    if (!IsInRange(win, xx, yy, row1, col1, row2, col2)) continue;
                    x = xx;
                    y = yy;
                    return true;
                }
            }
            return false;
        }

        static int FindBoundary(Func<int, bool> inside, int start, int dir, int limit)
        {
            int lastInside = start;
            int cursor = start;
            for (int i = 0; i < 80; i++)
            {
                int next = cursor + dir * 4;
                if (dir < 0 && next < limit) break;
                if (dir > 0 && next > limit) break;
                if (!inside(next))
                {
                    int best = lastInside;
                    for (int p = lastInside + dir; p != next + dir; p += dir)
                    {
                        if (inside(p)) best = p;
                        else break;
                    }
                    return dir > 0 ? best + 1 : best;
                }
                cursor = next;
                lastInside = next;
            }
            return lastInside;
        }

        static bool IsInRange(ExcelWindow win, int x, int y, int row1, int col1, int row2, int col2)
        {
            return TryCellAt(win, x, y, out int row, out int col)
                   && row >= row1 && row <= row2
                   && col >= col1 && col <= col2;
        }

        static bool TryCellAt(ExcelWindow win, int x, int y, out int row, out int col)
        {
            row = 0;
            col = 0;
            object obj = null;
            try
            {
                obj = win.RangeFromPoint(x, y);
                var cell = obj as ExcelRange;
                if (cell == null) return false;
                row = cell.Row;
                col = cell.Column;
                return row > 0 && col > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (obj != null && Marshal.IsComObject(obj))
                {
                    try { Marshal.ReleaseComObject(obj); } catch { }
                }
            }
        }

        /// <summary>
        /// VisibleRange の画面矩形とセル相対位置から推定（PointsToScreenPixels が不安定な環境向け）。
        /// </summary>
        static Rect? TryGetRangeScreenRectViaVisibleRange(
            ExcelApp excelApp, ExcelRange range, double winLeft, double winTop, double winW, double winH)
        {
            try
            {
                ExcelWindow win = excelApp.ActiveWindow;
                if (win == null) return null;
                ExcelRange visible = win.VisibleRange as ExcelRange;
                if (visible == null) return null;

                double vLeft = Convert.ToDouble(visible.Left);
                double vTop = Convert.ToDouble(visible.Top);
                double vW = Math.Max(1, Convert.ToDouble(visible.Width));
                double vH = Math.Max(1, Convert.ToDouble(visible.Height));

                int vx1 = win.PointsToScreenPixelsX((int)Math.Round(vLeft));
                int vy1 = win.PointsToScreenPixelsY((int)Math.Round(vTop));
                int vx2 = win.PointsToScreenPixelsX((int)Math.Round(vLeft + vW));
                int vy2 = win.PointsToScreenPixelsY((int)Math.Round(vTop + vH));

                double screenLeft = Math.Min(vx1, vx2);
                double screenTop = Math.Min(vy1, vy2);
                double screenW = Math.Max(20, Math.Abs(vx2 - vx1));
                double screenH = Math.Max(20, Math.Abs(vy2 - vy1));

                double rLeft = Convert.ToDouble(range.Left);
                double rTop = Convert.ToDouble(range.Top);
                double rW = Convert.ToDouble(range.Width);
                double rH = Convert.ToDouble(range.Height);

                double x = screenLeft + (rLeft - vLeft) / vW * screenW;
                double y = screenTop + (rTop - vTop) / vH * screenH;
                double w = rW / vW * screenW;
                double h = rH / vH * screenH;

                var rect = new Rect(x, y, Math.Max(20, w), Math.Max(20, h));
                rect.Intersect(new Rect(winLeft, winTop, winW, winH));
                if (rect.Width < 16 || rect.Height < 16) return null;
                return rect;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>オーバーレイ失敗時でも見えるよう、シート上にオレンジ枠シェイプを置く。</summary>
        public static void ApplyNativeTableHighlight(ExcelApp excelApp)
        {
            ClearNativeTableHighlight(excelApp);
            if (excelApp == null) return;
            try
            {
                ExcelWorksheet ws = excelApp.ActiveSheet as ExcelWorksheet;
                ExcelRange range = TryGetVocabTableRange(excelApp);
                if (ws == null || range == null) return;

                float l = Convert.ToSingle(range.Left);
                float t = Convert.ToSingle(range.Top);
                float w = Convert.ToSingle(range.Width);
                float h = Convert.ToSingle(range.Height);
                if (w < 4 || h < 4) return;

                // Office.Core 非依存: セル上に枠（クリックを奪わない BorderAround を優先）
                try
                {
                    range.BorderAround(
                        Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous,
                        Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick,
                        Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic,
                        OrangeRgb);
                    return;
                }
                catch { }

                dynamic shapes = ws.Shapes;
                dynamic shape = shapes.AddShape(1, l, t, w, h); // msoShapeRectangle
                shape.Name = NativeHighlightShapeName;
                try { shape.Placement = 1; } catch { } // xlMoveAndSize
                try { shape.ZOrder(1); } catch { } // msoSendToBack — セルクリックを優先
                try { shape.Fill.Visible = 0; } catch { try { shape.Fill.Transparency = 1.0; } catch { } }
                try
                {
                    shape.Line.Visible = -1;
                    shape.Line.ForeColor.RGB = OrangeRgb;
                    shape.Line.Weight = 3.5;
                }
                catch { }
            }
            catch
            {
                // ネイティブ強調は補助。失敗しても続行
            }
        }

        public static void ClearNativeTableHighlight(ExcelApp excelApp)
        {
            if (excelApp == null) return;
            try
            {
                ExcelWorksheet ws = excelApp.ActiveSheet as ExcelWorksheet;
                if (ws == null) return;

                try
                {
                    dynamic shapes = ws.Shapes;
                    dynamic shape = shapes.Item(NativeHighlightShapeName);
                    shape.Delete();
                }
                catch { }

                // BorderAround のオレンジ枠を外す（テーブル書式は残る）
                try
                {
                    ExcelRange range = TryGetVocabTableRange(excelApp);
                    if (range == null) return;
                    var edges = new[]
                    {
                        Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft,
                        Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop,
                        Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom,
                        Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight,
                    };
                    foreach (var edge in edges)
                    {
                        try
                        {
                            var b = range.Borders[edge];
                            int rgb = 0;
                            try { rgb = Convert.ToInt32(b.Color); } catch { }
                            if (rgb == OrangeRgb)
                                b.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                        }
                        catch { }
                    }
                }
                catch { }
            }
            catch { }
        }

        public static bool IsTableCurrentlySelected(ExcelApp excelApp)
        {
            if (excelApp == null) return false;
            try
            {
                var sel = excelApp.Selection as ExcelRange;
                if (sel == null) return false;
                try
                {
                    if (sel.ListObject != null) return true;
                }
                catch { }

                ExcelWorksheet ws = excelApp.ActiveSheet as ExcelWorksheet;
                if (ws == null) return false;
                ExcelRange table = ws.Range[VocabTableAddress];
                ExcelRange inter = excelApp.Intersect(sel, table);
                return inter != null;
            }
            catch
            {
                return false;
            }
        }

        public static bool IsChartCurrentlySelected(ExcelApp excelApp)
        {
            if (excelApp == null) return false;
            try
            {
                object sel = excelApp.Selection;
                if (sel == null) return false;
                if (sel is Microsoft.Office.Interop.Excel.Chart) return true;
                if (sel is Microsoft.Office.Interop.Excel.ChartObject) return true;
                string name = "";
                try { name = sel.GetType().Name ?? ""; } catch { }
                return name.IndexOf("Chart", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }
    }
}
