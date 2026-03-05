using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using Libraries;

namespace Libraries.Group1
{
    public class PowerPointChecker1_11
    {
        /// <summary>11-1: スライドサイズが 16:10（画面に合わせる）。</summary>
        public bool CheckTask_1_11_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                PageSetup pageSetup = null;
                try
                {
                    pageSetup = pres.PageSetup;
                    if (pageSetup == null) return false;
                    try
                    {
                        float w = (float)pageSetup.SlideWidth;
                        float h = (float)pageSetup.SlideHeight;
                        if (w <= 0) return false;
                        float ratio = h / w;
                        // 16:10 = 1.6
                        return Math.Abs(ratio - 1.6f) < 0.02f;
                    }
                    catch { return false; }
                }
                finally
                {
                    if (pageSetup != null) { try { Marshal.ReleaseComObject(pageSetup); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>11-2: スライド1のセクション名が「タイトル」。</summary>
        public bool CheckTask_1_11_02()
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
                    SectionProperties sectionProps = null;
                    try
                    {
                        sectionProps = pres.SectionProperties;
                        if (sectionProps == null) return false;
                        try
                        {
                            int sectionIndex = (int)slide.sectionIndex;
                            if (sectionIndex < 1) return false;
                            string name = null;
                            try { name = sectionProps.Name(sectionIndex) ?? ""; } catch { return false; }
                            return name.IndexOf("タイトル", StringComparison.OrdinalIgnoreCase) >= 0;
                        }
                        catch { return false; }
                    }
                    finally
                    {
                        if (sectionProps != null) { try { Marshal.ReleaseComObject(sectionProps); } catch { } }
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

        /// <summary>11-3: スライド3の箇条書きテキストボックスに塗りつぶし・枠線1.5pt。</summary>
        public bool CheckTask_1_11_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 3);
                    if (slide == null) return false;
                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
                        if (shapes == null) return false;
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (sh.HasTextFrame != MsoTriState.msoTrue) continue;
                                string text = null;
                                try
                                {
                                    var tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                    if (tf == null) continue;
                                    text = tf.TextRange?.Text ?? "";
                                }
                                catch { continue; }
                                if (string.IsNullOrEmpty(text) || text.IndexOf("•", StringComparison.Ordinal) < 0 && text.IndexOf("・", StringComparison.Ordinal) < 0) continue;
                                try
                                {
                                    Microsoft.Office.Interop.PowerPoint.FillFormat fill = sh.Fill;
                                    if (fill == null) continue;
                                    try
                                    {
                                        if (fill.Visible != MsoTriState.msoTrue) continue;
                                        Microsoft.Office.Interop.PowerPoint.ColorFormat cf = fill.ForeColor;
                                        if (cf == null) continue;
                                        try
                                        {
                                            bool accent1 = (cf.ObjectThemeColor == MsoThemeColorIndex.msoThemeColorAccent1);
                                            Microsoft.Office.Interop.PowerPoint.LineFormat line = sh.Line;
                                            if (line == null) continue;
                                            try
                                            {
                                                float weight = (float)line.Weight;
                                                return accent1 && Math.Abs(weight - 1.5f) < 0.2f;
                                            }
                                            finally { if (line != null) { try { Marshal.ReleaseComObject(line); } catch { } } }
                                        }
                                        finally { if (cf != null) { try { Marshal.ReleaseComObject(cf); } catch { } } }
                                    }
                                    finally { if (fill != null) { try { Marshal.ReleaseComObject(fill); } catch { } } }
                                }
                                catch { continue; }
                            }
                            finally { if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } } }
                        }
                        return false;
                    }
                    finally { if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } } }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>11-4: 配布資料マスターで日付削除・フッター「四季のうつろい」。</summary>
        public bool CheckTask_1_11_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Master handoutMaster = null;
                try
                {
                    handoutMaster = pres.HandoutMaster;
                    if (handoutMaster == null) return false;
                    try
                    {
                        HeadersFooters hf = null;
                        try
                        {
                            hf = handoutMaster.HeadersFooters;
                            if (hf == null) return false;
                            bool dateVisible = hf.DateAndTime.Visible == MsoTriState.msoTrue;
                            string footer = hf.Footer.Text ?? "";
                            return !dateVisible && footer.IndexOf("四季のうつろい", StringComparison.OrdinalIgnoreCase) >= 0;
                        }
                        finally { if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } } }
                    }
                    catch { return false; }
                }
                finally { if (handoutMaster != null) { try { Marshal.ReleaseComObject(handoutMaster); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>11-5: スライド5のアイコンに「青」塗りつぶし。</summary>
        public bool CheckTask_1_11_05()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    if (slide == null) return false;
                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
                        if (shapes == null) return false;
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                Microsoft.Office.Interop.PowerPoint.FillFormat fill = null;
                                try
                                {
                                    fill = sh.Fill;
                                    if (fill == null || fill.Visible != MsoTriState.msoTrue) continue;
                                    Microsoft.Office.Interop.PowerPoint.ColorFormat cf = fill.ForeColor;
                                    if (cf == null) continue;
                                    try
                                    {
                                        var otheme = cf.ObjectThemeColor;
                                        if (otheme == MsoThemeColorIndex.msoThemeColorAccent1 || otheme == MsoThemeColorIndex.msoThemeColorAccent2) return true;
                                        int rgb = (int)cf.RGB;
                                        int b = rgb & 0xFF; int g = (rgb >> 8) & 0xFF; int r = (rgb >> 16) & 0xFF;
                                        if (b > 200 && r < 100 && g < 100) return true;
                                    }
                                    catch { continue; }
                                }
                                finally { if (fill != null) { try { Marshal.ReleaseComObject(fill); } catch { } } }
                            }
                            finally { if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } } }
                        }
                        return false;
                    }
                    finally { if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } } }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>11-6: スライド2の箇条書きテキストボックスを上下中央揃え。</summary>
        public bool CheckTask_1_11_06()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 2);
                    if (slide == null) return false;
                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
                        if (shapes == null) return false;
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (sh.HasTextFrame != MsoTriState.msoTrue) continue;
                                try
                                {
                                    dynamic tf2 = sh.TextFrame2;
                                    if (tf2 == null) continue;
                                    try
                                    {
                                        int anchor = (int)tf2.VerticalAnchor;
                                        return anchor == (int)MsoVerticalAnchor.msoAnchorMiddle;
                                    }
                                    catch { continue; }
                                }
                                catch { continue; }
                            }
                            finally { if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } } }
                        }
                        return false;
                    }
                    finally { if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } } }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
        /// <summary>11-7: ノートで全スライド3部・部単位で印刷。COM の PrintOptions または VSTO ログの印刷記録で判定。</summary>
        public bool CheckTask_1_11_07()
        {
            if (PPLogReader.HasTask11_7PrintExecuted())
                return true;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                PrintOptions po = null;
                try
                {
                    po = pres.PrintOptions;
                    if (po == null) return false;
                    try
                    {
                        if (po.OutputType != PpPrintOutputType.ppPrintOutputNotesPages) return false;
                        if (po.NumberOfCopies != 3) return false;
                        return po.Collate == MsoTriState.msoTrue;
                    }
                    catch { return false; }
                }
                finally { if (po != null) { try { Marshal.ReleaseComObject(po); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
    }
}
