using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_3
    {
        public bool CheckTask_1_3_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                PptShape saShape = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    if (slide == null) return false;
                    saShape = PowerPointCheckerCommon.FindSmartArtShape(slide);
                    if (saShape == null) return false;
                    SmartArt smartArt = null;
                    try
                    {
                        smartArt = saShape.SmartArt;
                        if (smartArt == null) return false;
                        SmartArtNodes nodes = null;
                        try
                        {
                            nodes = smartArt.AllNodes;
                            if (nodes == null) return false;
                            var sb = new StringBuilder();
                            int nCount = nodes.Count;
                            for (int j = 1; j <= nCount; j++)
                            {
                                SmartArtNode node = null;
                                try
                                {
                                    node = nodes[j];
                                    if (node != null)
                                    {
                                        try
                                        {
                                            var tf2 = node.TextFrame2;
                                            if (tf2 != null && tf2.TextRange != null)
                                            {
                                                string t = tf2.TextRange.Text ?? "";
                                                sb.Append(t);
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                finally
                                {
                                    if (node != null) { try { Marshal.ReleaseComObject(node); } catch { } }
                                }
                            }
                            string allText = sb.ToString();
                            return allText.IndexOf("1F受付", StringComparison.OrdinalIgnoreCase) >= 0
                                && allText.IndexOf("面接室", StringComparison.OrdinalIgnoreCase) >= 0;
                        }
                        finally
                        {
                            if (nodes != null) { try { Marshal.ReleaseComObject(nodes); } catch { } }
                        }
                    }
                    finally
                    {
                        if (smartArt != null) { try { Marshal.ReleaseComObject(smartArt); } catch { } }
                    }
                }
                finally
                {
                    if (saShape != null) { try { Marshal.ReleaseComObject(saShape); } catch { } }
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_3_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                dynamic app = null;
                try
                {
                    app = pres.Application;
                    if (app == null) return false;
                }
                catch { return false; }
                Slide slide = null;
                PptShape saShape = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    if (slide == null) return false;
                    saShape = PowerPointCheckerCommon.FindSmartArtShape(slide);
                    if (saShape == null) return false;
                    SmartArt smartArt = null;
                    try
                    {
                        smartArt = saShape.SmartArt;
                        if (smartArt == null) return false;
                        try
                        {
                            var appliedColor = smartArt.Color;
                            if (appliedColor == null) return false;
                            try
                            {
                                string appliedId = appliedColor.Id ?? "";
                                string appliedName = appliedColor.Name ?? "";
                                if (appliedId.IndexOf("accent5", StringComparison.OrdinalIgnoreCase) >= 0)
                                    return true;
                                if (appliedName.IndexOf("アクセント5", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    appliedName.IndexOf("アクセント 5", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    appliedName.IndexOf("Accent 5", StringComparison.OrdinalIgnoreCase) >= 0)
                                    return true;
                                for (int idx = 1; idx <= 20; idx++)
                                {
                                    try
                                    {
                                        var style = app.SmartArtColors[idx];
                                        if (style == null) continue;
                                        string styleId = style.Id ?? "";
                                        string styleName = style.Name ?? "";
                                        if (styleName.IndexOf("アクセント5", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                            styleName.IndexOf("Accent 5", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            if (styleId == appliedId)
                                            {
                                                try { Marshal.ReleaseComObject(style); } catch { }
                                                return true;
                                            }
                                            try { Marshal.ReleaseComObject(style); } catch { }
                                            break;
                                        }
                                        try { Marshal.ReleaseComObject(style); } catch { }
                                    }
                                    catch { break; }
                                }
                                return false;
                            }
                            finally
                            {
                                try { Marshal.ReleaseComObject(appliedColor); } catch { }
                            }
                        }
                        catch { return false; }
                    }
                    finally
                    {
                        if (smartArt != null) { try { Marshal.ReleaseComObject(smartArt); } catch { } }
                    }
                }
                finally
                {
                    if (saShape != null) { try { Marshal.ReleaseComObject(saShape); } catch { } }
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>
        /// 3-3: スライド6の箇条書きを「縦方向カーブリスト」のSmartArtに変更したか判定。
        /// スライド6に SmartArt が1つあり、レイアウトが「縦方向カーブリスト」(VerticalCurvedList) であることを確認する。
        /// </summary>
        public bool CheckTask_1_3_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                PptShape saShape = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 6);
                    if (slide == null) return false;
                    saShape = PowerPointCheckerCommon.FindSmartArtShape(slide);
                    if (saShape == null) return false;
                    SmartArt smartArt = null;
                    try
                    {
                        smartArt = saShape.SmartArt;
                        if (smartArt == null) return false;
                        try
                        {
                            var layout = smartArt.Layout;
                            if (layout == null) return false;
                            try
                            {
                                string layoutId = layout.Id ?? "";
                                return layoutId.IndexOf("VerticalCurvedList", StringComparison.OrdinalIgnoreCase) >= 0;
                            }
                            finally
                            {
                                if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } }
                            }
                        }
                        catch { return false; }
                    }
                    finally
                    {
                        if (smartArt != null) { try { Marshal.ReleaseComObject(smartArt); } catch { } }
                    }
                }
                finally
                {
                    if (saShape != null) { try { Marshal.ReleaseComObject(saShape); } catch { } }
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>
        /// 3-4: スライド1に「この機能が使える！」「受験当日の流れ」のスライドズームを挿入し、文字より下に配置し重ならないようにしたか判定。
        /// 文字が入っているオブジェクトの下端（フッター・日付・スライド番号を除く）より下にズームが2つあり、重なっていないことを確認する。
        /// </summary>
        public bool CheckTask_1_3_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                    if (slide == null) return false;
                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
                        if (shapes == null) return false;

                        float slideHeight = 0f;
                        try
                        {
                            var pageSetup = pres.PageSetup;
                            if (pageSetup != null) slideHeight = (float)pageSetup.SlideHeight;
                        }
                        catch { }
                        if (slideHeight <= 0) slideHeight = 540f;

                        double textBottom = 0;
                        var candidateZooms = new List<Tuple<PptShape, float, float, float, float>>();
                        int count = shapes.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                float left = (float)sh.Left;
                                float top = (float)sh.Top;
                                float width = (float)sh.Width;
                                float height = (float)sh.Height;
                                int st = (int)sh.Type;

                                if (sh.HasTextFrame == MsoTriState.msoTrue)
                                {
                                    try
                                    {
                                        var tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                        if (tf.HasText == MsoTriState.msoTrue && tf.TextRange != null && !string.IsNullOrWhiteSpace(tf.TextRange.Text))
                                        {
                                            bool excludeFromText = (top >= slideHeight * 0.75f);
                                            if (!excludeFromText && (st == (int)MsoShapeType.msoPicture || st == 11 || st == 28 || st == 29 || st == 7 || st == 14))
                                                excludeFromText = true;
                                            if (!excludeFromText && st == (int)MsoShapeType.msoPlaceholder)
                                            {
                                                try
                                                {
                                                    var pf = sh.PlaceholderFormat;
                                                    if (pf != null)
                                                    {
                                                        try
                                                        {
                                                            PpPlaceholderType ppt = (PpPlaceholderType)pf.Type;
                                                            if (ppt == PpPlaceholderType.ppPlaceholderFooter ||
                                                                ppt == PpPlaceholderType.ppPlaceholderDate ||
                                                                ppt == PpPlaceholderType.ppPlaceholderSlideNumber)
                                                                excludeFromText = true;
                                                        }
                                                        finally { if (pf != null) { try { Marshal.ReleaseComObject(pf); } catch { } } }
                                                    }
                                                }
                                                catch { }
                                            }
                                            if (!excludeFromText)
                                            {
                                                double bottom = top + height;
                                                if (bottom > textBottom) textBottom = bottom;
                                            }
                                        }
                                    }
                                    catch { }
                                }
                                if (st == (int)MsoShapeType.msoPicture || st == 11 || st == 28 || st == 29
                                    || st == 7 || st == 14)
                                {
                                    candidateZooms.Add(Tuple.Create(sh, left, top, width, height));
                                    sh = null;
                                }
                            }
                            finally
                            {
                                if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                            }
                        }

                        const float belowTolerance = 5f;
                        float minTopForZoom = (float)textBottom + belowTolerance;
                        var zoomsBelowText = new List<Tuple<PptShape, float, float, float, float>>();
                        foreach (var t in candidateZooms)
                        {
                            if (t.Item3 >= minTopForZoom)
                                zoomsBelowText.Add(t);
                            else
                            {
                                try { Marshal.ReleaseComObject(t.Item1); } catch { }
                            }
                        }

                        if (zoomsBelowText.Count < 2)
                        {
                            foreach (var t in zoomsBelowText) { try { Marshal.ReleaseComObject(t.Item1); } catch { } }
                            return false;
                        }

                        bool noOverlap = true;
                        for (int a = 0; a < zoomsBelowText.Count && noOverlap; a++)
                        {
                            for (int b = a + 1; b < zoomsBelowText.Count && noOverlap; b++)
                            {
                                var ta = zoomsBelowText[a];
                                var tb = zoomsBelowText[b];
                                float la = ta.Item2, ra = ta.Item2 + ta.Item4, taTop = ta.Item3, ba = ta.Item3 + ta.Item5;
                                float lb = tb.Item2, rb = tb.Item2 + tb.Item4, tbTop = tb.Item3, bb = tb.Item3 + tb.Item5;
                                bool overlaps = !(ra <= lb || la >= rb || ba <= tbTop || taTop >= bb);
                                if (overlaps) noOverlap = false;
                            }
                        }

                        foreach (var t in zoomsBelowText) { try { Marshal.ReleaseComObject(t.Item1); } catch { } }
                        return noOverlap;
                    }
                    finally
                    {
                        if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
                    }
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
    }
}
