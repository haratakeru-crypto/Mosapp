using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using Libraries;

namespace Libraries.Group1
{
    /// <summary>プロジェクト7（P7-1〜P7-5）。Phase B でタスク単位に Legacy 移植。</summary>
    public class PowerPointChecker1_7
    {
        private const string ExpectedHyperlinkAddressP7_2 = "https://rabbitway.jp/";
        private const string ExpectedHyperlinkTextP7_2 = "情報学習支援";

        /// <summary>P7-1: スライド1にコメント「情報発信の責任を考える」（旧6-1、文言差分）。</summary>
        public bool CheckTask_1_7_01()
        {
            const string expectedComment = "情報発信の責任を考える";
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
                    Comments comments = null;
                    try
                    {
                        comments = slide.Comments;
                        if (comments == null) return false;
                        int count = comments.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            Comment cmt = null;
                            try
                            {
                                cmt = comments[i];
                                if (cmt == null) continue;
                                string text = null;
                                try { text = cmt.Text ?? ""; } catch { }
                                if (text.IndexOf(expectedComment, StringComparison.OrdinalIgnoreCase) >= 0)
                                    return true;
                            }
                            finally
                            {
                                if (cmt != null) { try { Marshal.ReleaseComObject(cmt); } catch { } }
                            }
                        }
                        return false;
                    }
                    finally
                    {
                        if (comments != null) { try { Marshal.ReleaseComObject(comments); } catch { } }
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

        /// <summary>P7-2: スライド1「情報学習支援」に https://rabbitway.jp/ のハイパーリンク（旧9-6、文言・URL差分）。</summary>
        public bool CheckTask_1_7_02()
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
                    PptShape shape = null;
                    try
                    {
                        shape = PowerPointCheckerCommon.FindShapeWithText(slide, ExpectedHyperlinkTextP7_2);
                        if (shape == null) return false;
                        TextFrame tf = null;
                        try
                        {
                            tf = shape.TextFrame;
                            if (tf == null) return false;
                            TextRange tr = null;
                            try
                            {
                                tr = tf.TextRange;
                                if (tr == null) return false;
                                if (TryHyperlinkOnTextRange(tr, ExpectedHyperlinkAddressP7_2))
                                    return true;
                                if (TryHyperlinkOnRunsContainingText(tr, ExpectedHyperlinkTextP7_2, ExpectedHyperlinkAddressP7_2))
                                    return true;
                                return false;
                            }
                            finally
                            {
                                if (tr != null) { try { Marshal.ReleaseComObject(tr); } catch { } }
                            }
                        }
                        finally
                        {
                            if (tf != null) { try { Marshal.ReleaseComObject(tf); } catch { } }
                        }
                    }
                    finally
                    {
                        if (shape != null) { try { Marshal.ReleaseComObject(shape); } catch { } }
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

        private static bool TryHyperlinkOnRunsContainingText(TextRange tr, string expectedText, string expectedAddress)
        {
            if (tr == null || string.IsNullOrEmpty(expectedText)) return false;

            int maxRunIndex = 256;
            try
            {
                int textLen = tr.Length;
                if (textLen > 0 && textLen < maxRunIndex)
                    maxRunIndex = textLen;
            }
            catch { }

            for (int r = 1; r <= maxRunIndex; r++)
            {
                TextRange run = null;
                try
                {
                    try { run = tr.Runs(r, 1); }
                    catch { break; }
                    if (run == null) break;

                    string runText = null;
                    try { runText = run.Text ?? ""; } catch { runText = ""; }
                    if (string.IsNullOrEmpty(runText)) break;

                    if (runText.IndexOf(expectedText, StringComparison.OrdinalIgnoreCase) >= 0
                        && TryHyperlinkOnTextRange(run, expectedAddress))
                        return true;
                }
                finally
                {
                    if (run != null) { try { Marshal.ReleaseComObject(run); } catch { } }
                }
            }
            return false;
        }

        private static bool TryHyperlinkOnTextRange(TextRange tr, string expectedAddress)
        {
            if (tr == null) return false;
            ActionSettings acts = null;
            try
            {
                acts = tr.ActionSettings;
                if (acts == null) return false;
                ActionSetting act = null;
                try
                {
                    act = acts[PpMouseActivation.ppMouseClick];
                    if (act == null) return false;
                    if (act.Action != PpActionType.ppActionHyperlink) return false;
                    Hyperlink hyp = null;
                    try
                    {
                        hyp = act.Hyperlink;
                        if (hyp == null) return false;
                        string addr = null;
                        try { addr = hyp.Address ?? ""; } catch { addr = ""; }
                        return HyperlinkAddressMatches(addr, expectedAddress);
                    }
                    finally
                    {
                        if (hyp != null) { try { Marshal.ReleaseComObject(hyp); } catch { } }
                    }
                }
                finally
                {
                    if (act != null) { try { Marshal.ReleaseComObject(act); } catch { } }
                }
            }
            finally
            {
                if (acts != null) { try { Marshal.ReleaseComObject(acts); } catch { } }
            }
        }

        private static bool HyperlinkAddressMatches(string address, string expected)
        {
            return string.Equals((address ?? "").Trim(), (expected ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>P7-3: スライド6の後ろに文書「まとめ」のアウトライン挿入。7枚目に「まとめ」を含む（旧7-3、文言・位置差分）。</summary>
        public bool CheckTask_1_7_03()
        {
            const string expectedOutlineText = "まとめ";
            const int expectedSlideNumber = 7;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null || slides.Count < expectedSlideNumber)
                        return false;
                }
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }

                Slide slide7 = null;
                try
                {
                    slide7 = PowerPointCheckerCommon.GetSlideByNumber(pres, expectedSlideNumber);
                    if (slide7 == null) return false;
                    PptShape sh = null;
                    try
                    {
                        sh = PowerPointCheckerCommon.FindShapeWithText(slide7, expectedOutlineText);
                        return sh != null;
                    }
                    finally { if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } } }
                }
                finally { if (slide7 != null) { try { Marshal.ReleaseComObject(slide7); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P7-4: フッターにスライド番号と rabbitway.jp。スライド2〜4必須・タイトル非表示（旧9-4、文言差分）。</summary>
        public bool CheckTask_1_7_04()
        {
            const string expectedFooterText = "rabbitway.jp";
            const int requiredLastSlide = 4;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null || slides.Count < requiredLastSlide) return false;

                    bool isSkipOnTitleEnabled = false;
                    try
                    {
                        if (pres.SlideMaster.HeadersFooters.DisplayOnTitleSlide == Microsoft.Office.Core.MsoTriState.msoFalse)
                            isSkipOnTitleEnabled = true;
                    }
                    catch { }

                    Slide slide1 = null;
                    try
                    {
                        slide1 = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                        if (slide1 != null)
                        {
                            HeadersFooters hf1 = null;
                            try
                            {
                                hf1 = slide1.HeadersFooters;
                                if (hf1 != null)
                                {
                                    try
                                    {
                                        if (hf1.DisplayOnTitleSlide == Microsoft.Office.Core.MsoTriState.msoFalse)
                                            isSkipOnTitleEnabled = true;
                                    }
                                    catch { }
                                    bool footerVis = false;
                                    try { footerVis = (hf1.Footer.Visible == Microsoft.Office.Core.MsoTriState.msoTrue); } catch { }
                                    if (!footerVis) isSkipOnTitleEnabled = true;
                                }
                            }
                            finally { if (hf1 != null) { try { Marshal.ReleaseComObject(hf1); } catch { } } }
                        }
                    }
                    finally { if (slide1 != null) { try { Marshal.ReleaseComObject(slide1); } catch { } } }

                    if (!isSkipOnTitleEnabled) return false;

                    for (int slideNum = 2; slideNum <= requiredLastSlide; slideNum++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
                            if (slide == null || !SlideHasP7_4Footer(slide, expectedFooterText))
                                return false;
                        }
                        finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
                    }

                    return true;
                }
                finally { if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        private static bool SlideHasP7_4Footer(Slide slide, string expectedFooterText)
        {
            if (slide == null || string.IsNullOrEmpty(expectedFooterText)) return false;
            HeadersFooters hf = null;
            try
            {
                hf = slide.HeadersFooters;
                if (hf == null) return false;

                bool footerVis = false;
                try { footerVis = (hf.Footer.Visible == Microsoft.Office.Core.MsoTriState.msoTrue); } catch { }
                if (!footerVis) return false;

                bool snVis = false;
                try { snVis = (hf.SlideNumber.Visible == Microsoft.Office.Core.MsoTriState.msoTrue); } catch { }
                if (!snVis) return false;

                string text = "";
                try { text = hf.Footer?.Text ?? ""; } catch { }
                return text.IndexOf(expectedFooterText, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch { return false; }
            finally { if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } } }
        }

        /// <summary>P7-5: スライド5・6のフッターにのみ「参考事例」（旧9-5、対象2枚・文言差分）。</summary>
        public bool CheckTask_1_7_05()
        {
            const string expectedFooterText = "参考事例";
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slide slide5 = null;
                Slide slide6 = null;
                try
                {
                    slide5 = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    slide6 = PowerPointCheckerCommon.GetSlideByNumber(pres, 6);
                    if (slide5 == null || slide6 == null) return false;

                    if (!SlideFooterContainsText(slide5, expectedFooterText)) return false;
                    if (!SlideFooterContainsText(slide6, expectedFooterText)) return false;
                }
                finally
                {
                    if (slide5 != null) { try { Marshal.ReleaseComObject(slide5); } catch { } }
                    if (slide6 != null) { try { Marshal.ReleaseComObject(slide6); } catch { } }
                }

                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null) return true;
                    for (int i = 1; i <= slides.Count; i++)
                    {
                        if (i == 5 || i == 6) continue;
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            if (SlideFooterContainsText(slide, expectedFooterText))
                                return false;
                        }
                        finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
                    }
                    return true;
                }
                finally { if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        private static bool SlideFooterContainsText(Slide slide, string searchText)
        {
            if (slide == null || string.IsNullOrEmpty(searchText)) return false;
            HeadersFooters hf = null;
            try
            {
                hf = slide.HeadersFooters;
                if (hf == null) return false;
                string footer = "";
                try { footer = hf.Footer?.Text ?? ""; } catch { }
                return footer.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch { return false; }
            finally { if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } } }
        }
    }
}
