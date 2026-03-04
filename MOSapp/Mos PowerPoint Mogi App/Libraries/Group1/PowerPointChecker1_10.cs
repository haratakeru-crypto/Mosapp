using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using Libraries;

namespace Libraries.Group1
{
    public class PowerPointChecker1_10
    {
        /// <summary>10-1: ドキュメント検査実行（間接: コメントが削除されているか等）。</summary>
        public bool CheckTask_1_10_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                int totalComments = 0;
                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null) return false;
                    for (int i = 1; i <= slides.Count; i++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            Comments comments = null;
                            try
                            {
                                comments = slide.Comments;
                                if (comments != null) totalComments += comments.Count;
                            }
                            finally { if (comments != null) { try { Marshal.ReleaseComObject(comments); } catch { } } }
                        }
                        finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
                    }
                    return totalComments == 0;
                }
                finally { if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>10-2: 目的別スライドショー「教育」が存在するか。</summary>
        public bool CheckTask_1_10_02()
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
                    NamedSlideShows namedShows = null;
                    try
                    {
                        namedShows = ssSettings.NamedSlideShows;
                        if (namedShows == null) return false;
                        int count = namedShows.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            NamedSlideShow ns = null;
                            try
                            {
                                ns = namedShows[i];
                                if (ns == null) continue;
                                string name = null;
                                try { name = ns.Name ?? ""; } catch { }
                                if (name.IndexOf("教育", StringComparison.OrdinalIgnoreCase) >= 0)
                                    return true;
                            }
                            finally
                            {
                                if (ns != null) { try { Marshal.ReleaseComObject(ns); } catch { } }
                            }
                        }
                        return false;
                    }
                    finally
                    {
                        if (namedShows != null) { try { Marshal.ReleaseComObject(namedShows); } catch { } }
                    }
                }
                finally
                {
                    if (ssSettings != null) { try { Marshal.ReleaseComObject(ssSettings); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>10-3: スライド6の「理念まとめ」下プレースホルダーの文字間隔3pt。</summary>
        public bool CheckTask_1_10_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 6);
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
                                        dynamic tr2 = tf2.TextRange;
                                        if (tr2 == null) continue;
                                        float spacing = (float)tr2.Font.Spacing;
                                        if (Math.Abs(spacing - 3f) < 0.5f) return true;
                                    }
                                    catch { }
                                }
                                catch { }
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

        /// <summary>10-4: グレースケールは表示状態のため COM では検証不可。VSTO ログで記録されていれば true。</summary>
        public bool CheckTask_1_10_04()
        {
            return PPLogReader.HasTask10_4GrayscaleExecuted();
        }

        /// <summary>10-5: スライドマスターのテーマを「イオン」に。</summary>
        public bool CheckTask_1_10_05()
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
                    Slides slides = null;
                    try
                    {
                        slides = pres.Slides;
                        if (slides == null) return false;
                        SlideRange range = null;
                        try
                        {
                            range = slides.Range(new object[] { slide.SlideIndex });
                            if (range == null) return false;
                            Design design = null;
                            try
                            {
                                design = range.Design;
                                if (design == null) return false;
                                string name = null;
                                try { name = design.Name ?? ""; } catch { return false; }
                                return name.IndexOf("イオン", StringComparison.OrdinalIgnoreCase) >= 0;
                            }
                            finally { if (design != null) { try { Marshal.ReleaseComObject(design); } catch { } } }
                        }
                        finally { if (range != null) { try { Marshal.ReleaseComObject(range); } catch { } } }
                    }
                    finally { if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } } }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>10-6: 「タイトルとコンテンツ」レイアウトの背景グラフィック非表示。</summary>
        public bool CheckTask_1_10_06()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Master master = null;
                try
                {
                    master = pres.SlideMaster;
                    if (master == null) return false;
                    CustomLayouts layouts = null;
                    try
                    {
                        layouts = master.CustomLayouts;
                        if (layouts == null) return false;
                        for (int i = 1; i <= layouts.Count; i++)
                        {
                            CustomLayout cl = null;
                            try
                            {
                                cl = layouts[i];
                                if (cl == null) continue;
                                string name = null;
                                try { name = cl.Name ?? ""; } catch { continue; }
                                if (name.IndexOf("タイトルとコンテンツ", StringComparison.OrdinalIgnoreCase) < 0) continue;
                                try
                                {
                                    return cl.DisplayMasterShapes == MsoTriState.msoFalse;
                                }
                                catch { return false; }
                            }
                            finally { if (cl != null) { try { Marshal.ReleaseComObject(cl); } catch { } } }
                        }
                        return false;
                    }
                    finally { if (layouts != null) { try { Marshal.ReleaseComObject(layouts); } catch { } } }
                }
                finally { if (master != null) { try { Marshal.ReleaseComObject(master); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
        /// <summary>10-7: スライドマスターにレイアウト「画像付きスライド」が存在するか。</summary>
        public bool CheckTask_1_10_07()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Master master = null;
                try
                {
                    master = pres.SlideMaster;
                    if (master == null) return false;
                    CustomLayouts layouts = null;
                    try
                    {
                        layouts = master.CustomLayouts;
                        if (layouts == null) return false;
                        for (int i = 1; i <= layouts.Count; i++)
                        {
                            CustomLayout cl = null;
                            try
                            {
                                cl = layouts[i];
                                if (cl == null) continue;
                                string name = null;
                                try { name = cl.Name ?? ""; } catch { continue; }
                                if (name.IndexOf("画像付きスライド", StringComparison.OrdinalIgnoreCase) >= 0)
                                    return true;
                            }
                            finally { if (cl != null) { try { Marshal.ReleaseComObject(cl); } catch { } } }
                        }
                        return false;
                    }
                    finally { if (layouts != null) { try { Marshal.ReleaseComObject(layouts); } catch { } } }
                }
                finally { if (master != null) { try { Marshal.ReleaseComObject(master); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
    }
}
