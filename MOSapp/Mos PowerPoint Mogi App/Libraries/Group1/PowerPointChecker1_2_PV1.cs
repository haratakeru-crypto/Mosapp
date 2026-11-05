using System;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using PptGroupShapes = Microsoft.Office.Interop.PowerPoint.GroupShapes;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題セット1 プロジェクト2（木陰工房）。最終状態のみで判定する。
    /// </summary>
    public class PowerPointChecker1_2_PV1
    {
        private const string SlideTitleWork = "作業は、けがきの次に加工です";
        private const string SlideTitleLearn = "学び方は、工房、自宅、動画の3つです";
        private const string SlideTitleTools = "道具を持っている人は、86人です";
        private const string SmartArtLayoutIdHint = "hProcess9";
        private const string SmartArtLayoutNameHint = "矢印と長方形のプロセス";
        private const int PpEffectPushLeft = 3853;
        private const int PpEffectPushRight = 3854;
        private const int XlBarClustered = 57;
        private const int ExpectedInitialSlideCount = 6;

        public bool CheckTask_1_2_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null) return false;
                    if (slides.Count != ExpectedInitialSlideCount + 1)
                        return false;

                    int workIndex = FindSlideIndexByTitle(slides, SlideTitleWork);
                    if (workIndex < 1 || workIndex >= slides.Count)
                        return false;

                    return IsTableLayoutSlide(slides, workIndex + 1);
                }
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        public bool CheckTask_1_2_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null || slides.Count < 1) return false;
                    for (int i = 1; i <= slides.Count; i++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            int effectVal = (int)slide.SlideShowTransition.EntryEffect;
                            // 日本語UIでは「右から」が 3853 になる環境がある（Legacy 2-1 と同じ）。
                            if (effectVal != PpEffectPushLeft && effectVal != PpEffectPushRight)
                                return false;
                        }
                        catch { return false; }
                        finally
                        {
                            if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                        }
                    }
                    return true;
                }
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        public bool CheckTask_1_2_03()
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
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleWork);
                    if (slide == null) return false;
                    saShape = PowerPointCheckerCommon.FindSmartArtShape(slide);
                    if (saShape == null) return false;

                    SmartArt smartArt = null;
                    try
                    {
                        smartArt = saShape.SmartArt;
                        if (smartArt == null) return false;

                        SmartArtLayout layout = null;
                        try
                        {
                            layout = smartArt.Layout;
                            if (layout == null) return false;
                            string layoutId = "";
                            string layoutName = "";
                            try { layoutId = layout.Id ?? ""; } catch { }
                            try { layoutName = layout.Name ?? ""; } catch { }
                            if (!IsArrowRectangleProcessLayout(layoutId, layoutName))
                                return false;
                        }
                        finally
                        {
                            if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } }
                        }

                        string allText = CollectSmartArtText(smartArt);
                        if (allText.IndexOf("けがき", StringComparison.OrdinalIgnoreCase) < 0)
                            return false;
                        if (allText.IndexOf("加工", StringComparison.OrdinalIgnoreCase) < 0)
                            return false;

                        int filledNodes = CountFilledSmartArtNodes(smartArt);
                        return filledNodes == 2;
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
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        public bool CheckTask_1_2_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleLearn);
                    if (slide == null) return false;
                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
                        if (shapes == null) return false;
                        int count = shapes.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (IsThreeShapeGroup(sh))
                                    return true;
                            }
                            finally
                            {
                                if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                            }
                        }
                        return false;
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
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        public bool CheckTask_1_2_05()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                SlideShowSettings ssSettings = null;
                try
                {
                    ssSettings = pres.SlideShowSettings;
                    if (ssSettings == null) return false;
                    return ssSettings.ShowType == PpSlideShowType.ppShowTypeKiosk;
                }
                catch { return false; }
                finally
                {
                    if (ssSettings != null) { try { Marshal.ReleaseComObject(ssSettings); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        public bool CheckTask_1_2_06()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleTools);
                    if (slide == null) return false;
                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
                        if (shapes == null) return false;
                        int count = shapes.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (sh.HasChart != MsoTriState.msoTrue) continue;
                                Chart chart = null;
                                try
                                {
                                    chart = sh.Chart;
                                    if (chart == null) continue;
                                    int ct;
                                    try { ct = (int)chart.ChartType; }
                                    catch { continue; }
                                    // ChartType のみ判定する。SeriesCollection / ChartData / XValues は
                                    // 埋め込み Excel（OLE）を起動し、採点後に閉じられなくなることがある。
                                    if (ct == XlBarClustered)
                                        return true;
                                }
                                finally
                                {
                                    if (chart != null) { try { Marshal.ReleaseComObject(chart); } catch { } }
                                }
                            }
                            finally
                            {
                                if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                            }
                        }
                        return false;
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
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        private static bool IsThreeShapeGroup(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                if (sh.Type != MsoShapeType.msoGroup) return false;
            }
            catch { return false; }

            PptGroupShapes group = null;
            try
            {
                group = sh.GroupItems;
                return group != null && group.Count == 3;
            }
            catch { return false; }
            finally
            {
                if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
            }
        }

        private static bool IsArrowRectangleProcessLayout(string layoutId, string layoutName)
        {
            if (!string.IsNullOrEmpty(layoutId) &&
                layoutId.IndexOf(SmartArtLayoutIdHint, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            if (!string.IsNullOrEmpty(layoutName) &&
                layoutName.IndexOf(SmartArtLayoutNameHint, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            if (!string.IsNullOrEmpty(layoutName) &&
                layoutName.IndexOf("Arrow", StringComparison.OrdinalIgnoreCase) >= 0 &&
                layoutName.IndexOf("Process", StringComparison.OrdinalIgnoreCase) >= 0 &&
                layoutName.IndexOf("Rectangle", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            return false;
        }

        private static string CollectSmartArtText(SmartArt smartArt)
        {
            var sb = new StringBuilder();
            SmartArtNodes nodes = null;
            try
            {
                nodes = smartArt.AllNodes;
                if (nodes == null) return "";
                int nCount = nodes.Count;
                for (int j = 1; j <= nCount; j++)
                {
                    SmartArtNode node = null;
                    try
                    {
                        node = nodes[j];
                        if (node == null) continue;
                        try
                        {
                            var tf2 = node.TextFrame2;
                            if (tf2 != null && tf2.TextRange != null)
                                sb.Append(tf2.TextRange.Text ?? "");
                        }
                        catch { }
                    }
                    finally
                    {
                        if (node != null) { try { Marshal.ReleaseComObject(node); } catch { } }
                    }
                }
            }
            finally
            {
                if (nodes != null) { try { Marshal.ReleaseComObject(nodes); } catch { } }
            }
            return sb.ToString();
        }

        private static int CountFilledSmartArtNodes(SmartArt smartArt)
        {
            int filled = 0;
            SmartArtNodes nodes = null;
            try
            {
                nodes = smartArt.AllNodes;
                if (nodes == null) return 0;
                int nCount = nodes.Count;
                for (int j = 1; j <= nCount; j++)
                {
                    SmartArtNode node = null;
                    try
                    {
                        node = nodes[j];
                        if (node == null) continue;
                        string text = "";
                        try
                        {
                            var tf2 = node.TextFrame2;
                            if (tf2 != null && tf2.TextRange != null)
                                text = (tf2.TextRange.Text ?? "").Trim();
                        }
                        catch { }
                        if (!string.IsNullOrEmpty(text))
                            filled++;
                    }
                    finally
                    {
                        if (node != null) { try { Marshal.ReleaseComObject(node); } catch { } }
                    }
                }
            }
            finally
            {
                if (nodes != null) { try { Marshal.ReleaseComObject(nodes); } catch { } }
            }
            return filled;
        }

        private static int FindSlideIndexByTitle(Slides slides, string titlePart)
        {
            if (slides == null || string.IsNullOrEmpty(titlePart)) return -1;
            int count = slides.Count;
            for (int i = 1; i <= count; i++)
            {
                Slide slide = null;
                try
                {
                    slide = slides[i];
                    if (SlideContainsText(slide, titlePart))
                        return i;
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            return -1;
        }

        private static bool SlideContainsText(Slide slide, string titlePart)
        {
            if (slide == null || string.IsNullOrEmpty(titlePart)) return false;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return false;
                int sc = shapes.Count;
                for (int j = 1; j <= sc; j++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[j];
                        if (sh.HasTextFrame != MsoTriState.msoTrue) continue;
                        string text = null;
                        try
                        {
                            var pptTf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                            text = pptTf.TextRange.Text ?? "";
                        }
                        catch { continue; }
                        if (text.IndexOf(titlePart, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static bool IsTableLayoutSlide(Slides slides, int slideNumber)
        {
            if (slides == null || slideNumber < 1 || slideNumber > slides.Count)
                return false;
            Slide slide = null;
            CustomLayout layout = null;
            try
            {
                slide = slides[slideNumber];
                if (slide == null) return false;
                layout = slide.CustomLayout;
                if (layout == null) return false;
                string layoutName = "";
                try { layoutName = layout.Name ?? ""; } catch { return false; }
                return layoutName.IndexOf("表スライド", StringComparison.OrdinalIgnoreCase) >= 0
                    || layoutName.IndexOf("テーブルスライド", StringComparison.OrdinalIgnoreCase) >= 0
                    || layoutName.IndexOf("Table", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch { return false; }
            finally
            {
                if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } }
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }
    }
}
