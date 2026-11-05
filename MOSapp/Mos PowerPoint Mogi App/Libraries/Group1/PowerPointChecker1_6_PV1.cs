using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題セット1 プロジェクト6（ひがしの台所）。最終状態のみで判定する。
    /// ActivePresentation は ReleaseComObject しない（その場採点でプレゼンが閉じるのを防ぐ）。
    /// </summary>
    public class PowerPointChecker1_6_PV1
    {
        private const string SlideTitleTaste = "味は、出汁、醤油、柚子の順に足します";
        private const string SlideTitleKombu = "昆布は、沸騰の前に取り出します";
        private const string ExpectedFooterText = "手元に残す";
        private const string ExpectedHandoutFooter = "春の食卓";
        private const string ExpectedLayoutName = "タイトルとコンテンツ";
        private const float CmToPt = 72f / 2.54f;
        private const float ExpectedWidthCm = 33.12f;
        private const float ExpectedHeightCm = 25.04f;
        private const float SizeTolerancePt = 1.0f;
        private const float ExpectedTrimStartMs = 5000f;
        private const float ExpectedTrimEndMs = 10000f;
        private const float TrimToleranceMs = 500f;

        /// <summary>
        /// 類1-6-1: 指定2スライドのフッターのみ「手元に残す」。他スライドには出さない。
        /// </summary>
        public bool CheckTask_1_6_01()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;

                    Slide slideTaste = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleTaste);
                    Slide slideKombu = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleKombu);
                    try
                    {
                        if (slideTaste == null || slideKombu == null) return false;
                        if (!SlideHasFooterText(slideTaste, ExpectedFooterText)) return false;
                        if (!SlideHasFooterText(slideKombu, ExpectedFooterText)) return false;

                        int tasteId = slideTaste.SlideID;
                        int kombuId = slideKombu.SlideID;
                        Slides slides = pres.Slides;
                        if (slides == null) return false;
                        for (int i = 1; i <= slides.Count; i++)
                        {
                            Slide slide = null;
                            try
                            {
                                slide = slides[i];
                                if (slide == null) continue;
                                int id = slide.SlideID;
                                if (id == tasteId || id == kombuId) continue;
                                if (SlideHasFooterText(slide, ExpectedFooterText))
                                    return false;
                            }
                            finally
                            {
                                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                            }
                        }
                        return true;
                    }
                    finally
                    {
                        if (slideTaste != null) { try { Marshal.ReleaseComObject(slideTaste); } catch { } }
                        if (slideKombu != null) { try { Marshal.ReleaseComObject(slideKombu); } catch { } }
                    }
                }
                catch { return false; }
            }
        }

        /// <summary>
        /// 類1-6-2: スライドサイズ 幅33.12cm × 高さ25.04cm。
        /// （「サイズに合わせて調整」は最終サイズに反映されるが、COM で選択そのものは残らないため寸法で判定）
        /// </summary>
        public bool CheckTask_1_6_02()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    PageSetup pageSetup = pres.PageSetup;
                    if (pageSetup == null) return false;
                    float w = (float)pageSetup.SlideWidth;
                    float h = (float)pageSetup.SlideHeight;
                    float expectedW = ExpectedWidthCm * CmToPt;
                    float expectedH = ExpectedHeightCm * CmToPt;
                    return Math.Abs(w - expectedW) < SizeTolerancePt
                        && Math.Abs(h - expectedH) < SizeTolerancePt;
                }
                catch { return false; }
            }
        }

        /// <summary>類1-6-3: 動画を開始00:05・終了00:10にトリミング。</summary>
        public bool CheckTask_1_6_03()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleKombu);
                    if (slide == null) return false;
                    PptShape videoShape = null;
                    try
                    {
                        var status = PowerPointCheckerCommon.TryFindFirstVideoShape(slide, out videoShape);
                        if (status != PowerPointCheckerCommon.MediaScanStatus.Found || videoShape == null)
                            return false;

                        MediaFormat mf = null;
                        try
                        {
                            mf = videoShape.MediaFormat;
                            if (mf == null) return false;
                            float startPt = (float)mf.StartPoint;
                            float endPt = (float)mf.EndPoint;
                            return Math.Abs(startPt - ExpectedTrimStartMs) < TrimToleranceMs
                                && Math.Abs(endPt - ExpectedTrimEndMs) < TrimToleranceMs;
                        }
                        catch { return false; }
                        finally
                        {
                            if (mf != null) { try { Marshal.ReleaseComObject(mf); } catch { } }
                        }
                    }
                    finally
                    {
                        if (videoShape != null) { try { Marshal.ReleaseComObject(videoShape); } catch { } }
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        /// <summary>類1-6-4: 「タイトルとコンテンツ」レイアウトで背景グラフィック非表示。</summary>
        public bool CheckTask_1_6_04()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Master master = pres.SlideMaster;
                    if (master == null) return false;
                    CustomLayouts layouts = master.CustomLayouts;
                    if (layouts == null) return false;
                    for (int i = 1; i <= layouts.Count; i++)
                    {
                        CustomLayout cl = null;
                        try
                        {
                            cl = layouts[i];
                            if (cl == null) continue;
                            string name = "";
                            try { name = cl.Name ?? ""; } catch { continue; }
                            if (!IsTitleAndContentLayoutName(name))
                                continue;
                            return cl.DisplayMasterShapes == MsoTriState.msoFalse;
                        }
                        finally
                        {
                            if (cl != null) { try { Marshal.ReleaseComObject(cl); } catch { } }
                        }
                    }
                    return false;
                }
                catch { return false; }
            }
        }

        /// <summary>類1-6-5: 配布資料マスターで日付非表示・フッター「春の食卓」。</summary>
        public bool CheckTask_1_6_05()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Master handoutMaster = pres.HandoutMaster;
                    if (handoutMaster == null) return false;
                    HeadersFooters hf = null;
                    try
                    {
                        hf = handoutMaster.HeadersFooters;
                        if (hf == null) return false;
                        string footer = "";
                        try { footer = hf.Footer?.Text ?? ""; } catch { }
                        if (footer.IndexOf(ExpectedHandoutFooter, StringComparison.OrdinalIgnoreCase) < 0)
                            return false;
                        return !HandoutMasterHasVisibleDatePlaceholder(handoutMaster);
                    }
                    finally
                    {
                        if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } }
                    }
                }
                catch { return false; }
            }
        }

        private static bool SlideHasFooterText(Slide slide, string expectedFooterText)
        {
            if (slide == null || string.IsNullOrEmpty(expectedFooterText)) return false;
            HeadersFooters hf = null;
            try
            {
                hf = slide.HeadersFooters;
                if (hf == null) return false;

                bool footerVis = false;
                try { footerVis = (hf.Footer.Visible == MsoTriState.msoTrue); } catch { }
                if (!footerVis) return false;

                string text = "";
                try { text = hf.Footer?.Text ?? ""; } catch { }
                return text.IndexOf(expectedFooterText, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch { return false; }
            finally
            {
                if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } }
            }
        }

        private static bool IsTitleAndContentLayoutName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            if (name.IndexOf("タイトルとコンテンツ", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            if (name.IndexOf("Title and Content", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            if (name.IndexOf("Title And Content", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            return false;
        }

        private static bool HandoutMasterHasVisibleDatePlaceholder(Master handoutMaster)
        {
            if (handoutMaster == null) return false;
            PptShapes shapes = null;
            try
            {
                shapes = handoutMaster.Shapes;
                if (shapes == null) return false;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        if (sh == null) continue;
                        if (sh.Type != MsoShapeType.msoPlaceholder) continue;
                        PlaceholderFormat pf = null;
                        try
                        {
                            pf = sh.PlaceholderFormat;
                            if (pf == null) continue;
                            if ((PpPlaceholderType)pf.Type != PpPlaceholderType.ppPlaceholderDate)
                                continue;
                            if (sh.Visible == MsoTriState.msoTrue)
                                return true;
                        }
                        finally
                        {
                            if (pf != null) { try { Marshal.ReleaseComObject(pf); } catch { } }
                        }
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
    }
}
