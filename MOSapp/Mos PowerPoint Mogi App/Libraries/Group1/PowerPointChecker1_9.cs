using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_9
    {
        private const int XlBarClustered = 57;
        private const string ExpectedHyperlinkAddressTask9_6 = "https://www.jica.go.jp/activities/issues/natural_env/index.html";

        public bool CheckTask_1_9_01()
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
                                    try
                                    {
                                        int ct = (int)chart.ChartType;
                                        return ct == XlBarClustered;
                                    }
                                    finally
                                    {
                                        if (chart != null) { try { Marshal.ReleaseComObject(chart); } catch { } }
                                    }
                                }
                                catch { continue; }
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
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>9-2: スライド4の表を「中間スタイル1-アクセント4」に変更し、1行ずつ色が変わらないように（縞模様行OFF）。</summary>
        public bool CheckTask_1_9_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 4);
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
                                if (sh.HasTable != MsoTriState.msoTrue) continue;
                                Table table = null;
                                try
                                {
                                    table = sh.Table;
                                    if (table == null) continue;
                                    TableStyle ts = null;
                                    try
                                    {
                                        ts = table.Style;
                                        if (ts == null) continue;
                                        string styleName = "";
                                        try { styleName = ts.Name ?? ""; } catch { }
                                        // スペースあり・なしの両方に対応します
                                        bool hasAccent4 = styleName.Contains("アクセント 4") || styleName.Contains("Accent 4") || styleName.Contains("アクセント4");
                                        bool isBandedOff = !table.HorizBanding;
                                        return hasAccent4 && isBandedOff;
                                    }
                                    finally { if (ts != null) { try { Marshal.ReleaseComObject(ts); } catch { } } }
                                }
                                finally { if (table != null) { try { Marshal.ReleaseComObject(table); } catch { } } }
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
        private const int XlLabelPositionOutsideEnd = 2;

        /// <summary>9-3: スライド2のグラフのデータラベルを外側に追加。</summary>
        public bool CheckTask_1_9_03()
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
                                if (sh.HasChart != MsoTriState.msoTrue) continue;
                                Chart chart = null;
                                try
                                {
                                    chart = sh.Chart;
                                    if (chart == null) continue;
                                    try
                                    {
                                        var seriesColl = chart.SeriesCollection();
                                        if (seriesColl == null || seriesColl.Count < 1) continue;
                                        try
                                        {
                                            Series series = seriesColl.Item(1);
                                            if (series == null) continue;
                                            try
                                            {
                                                if (!series.HasDataLabels) continue;
                                                DataLabels dataLabels = series.DataLabels();
                                                if (dataLabels == null) continue;
                                                try
                                                {
                                                    int pos = (int)dataLabels.Position;
                                                    return pos == XlLabelPositionOutsideEnd;
                                                }
                                                finally { if (dataLabels != null) { try { Marshal.ReleaseComObject(dataLabels); } catch { } } }
                                            }
                                            finally { if (series != null) { try { Marshal.ReleaseComObject(series); } catch { } } }
                                        }
                                        finally { if (seriesColl != null) { try { Marshal.ReleaseComObject(seriesColl); } catch { } } }
                                    }
                                    finally { if (chart != null) { try { Marshal.ReleaseComObject(chart); } catch { } } }
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
        /// <summary>9-4: フッターに「www.MOS.jp」とページ番号を、タイトルスライド以外に追加。</summary>
        public bool CheckTask_1_9_04()
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
                    if (slides == null || slides.Count < 2) return false;

                    // 1. 「タイトルスライドに表示しない」設定の確認
                    // マスター設定、または個別のスライド設定のいずれかでオプションが有効かを確認します。
                    bool isSkipOnTitleEnabled = false;
                    try {
                        if (pres.SlideMaster.HeadersFooters.DisplayOnTitleSlide == MsoTriState.msoFalse) isSkipOnTitleEnabled = true;
                    } catch { }

                    // 2. スライドの内容をチェック
                    bool isOtherSlidesOk = false;
                    for (int i = 1; i <= slides.Count; i++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            HeadersFooters hf = null;
                            try
                            {
                                hf = slide.HeadersFooters;
                                if (hf == null) continue;

                                bool displayOnTitleIsFalse = false;
                                try { displayOnTitleIsFalse = (hf.DisplayOnTitleSlide == MsoTriState.msoFalse); } catch { }

                                bool footerVis = false;
                                try { footerVis = (hf.Footer.Visible == MsoTriState.msoTrue); } catch { }
                                
                                bool snVis = false;
                                try { snVis = (hf.SlideNumber.Visible == MsoTriState.msoTrue); } catch { }

                                // どこかのスライドで設定が確認できれば有効とみなす
                                if (displayOnTitleIsFalse) isSkipOnTitleEnabled = true;
                                // もしスライド1のフッターが非表示なら、設定は効いていると判断
                                if (i == 1 && !footerVis) isSkipOnTitleEnabled = true;

                                // スライド2以降でフッターが表示されている箇所があれば内容をチェック
                                if (i > 1 && footerVis)
                                {
                                    string text = "";
                                    try { text = hf.Footer.Text ?? ""; } catch { }
                                    bool isCorrectText = text.IndexOf("www.MOS.jp", StringComparison.OrdinalIgnoreCase) >= 0;
                                    
                                    // 9-5（5枚目）の変更が行われていても許容する
                                    if (i == 5 && text.IndexOf("集中的に", StringComparison.OrdinalIgnoreCase) >= 0) isCorrectText = true;

                                    if (isCorrectText && snVis)
                                    {
                                        isOtherSlidesOk = true;
                                    }
                                }
                            }
                            catch { }
                            finally { if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } } }
                        }
                        finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
                    }

                    // 設定が有効（またはスライド1で非表示）であり、かつ通常スライドの内容が正しければ合格
                    return isSkipOnTitleEnabled && isOtherSlidesOk;
                }
                finally { if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>9-5: 5枚目のスライドのフッターにのみ「集中的に」を挿入。</summary>
        public bool CheckTask_1_9_05()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide5 = null;
                try
                {
                    slide5 = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    if (slide5 == null) return false;
                    HeadersFooters hf5 = null;
                    try
                    {
                        hf5 = slide5.HeadersFooters;
                        if (hf5 == null) return false;
                        string footer5 = "";
                        try { footer5 = hf5.Footer?.Text ?? ""; } catch { }
                        if (footer5.IndexOf("集中的に", StringComparison.OrdinalIgnoreCase) < 0) return false;
                    }
                    finally { if (hf5 != null) { try { Marshal.ReleaseComObject(hf5); } catch { } } }
                }
                finally { if (slide5 != null) { try { Marshal.ReleaseComObject(slide5); } catch { } } }
                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null) return true;
                    for (int i = 1; i <= slides.Count; i++)
                    {
                        if (i == 5) continue;
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            HeadersFooters hf = null;
                            try
                            {
                                hf = slide.HeadersFooters;
                                if (hf == null) continue;
                                string footer = "";
                                try { footer = hf.Footer?.Text ?? ""; } catch { }
                                if (footer.IndexOf("集中的に", StringComparison.OrdinalIgnoreCase) >= 0) return false;
                            }
                            finally { if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } } }
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

        public bool CheckTask_1_9_06()
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
                        shape = PowerPointCheckerCommon.FindShapeWithText(slide, "お問い合わせ");
                        if (shape == null) return false;
                        try
                        {
                            Microsoft.Office.Interop.PowerPoint.TextFrame tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)shape.TextFrame;
                            if (tf == null) return false;
                            TextRange tr = null;
                            try
                            {
                                tr = tf.TextRange;
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
                                            if (hyp != null)
                                            {
                                                string addr = (hyp.Address ?? "").Trim();
                                                if (string.Equals(addr, ExpectedHyperlinkAddressTask9_6, StringComparison.OrdinalIgnoreCase))
                                                    return true;
                                            }
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
                                try
                                {
                                    int r = 1;
                                    while (true)
                                    {
                                        TextRange run = null;
                                        try
                                        {
                                            run = tr.Runs(r, 1);
                                            if (run == null) break;
                                            ActionSettings runActs = null;
                                            try
                                            {
                                                runActs = run.ActionSettings;
                                                if (runActs == null) { r++; continue; }
                                                ActionSetting runAct = null;
                                                try
                                                {
                                                    runAct = runActs[PpMouseActivation.ppMouseClick];
                                                    if (runAct != null && runAct.Action == PpActionType.ppActionHyperlink)
                                                    {
                                                        Hyperlink runHyp = null;
                                                        try
                                                        {
                                                            runHyp = runAct.Hyperlink;
                                                            if (runHyp != null)
                                                            {
                                                                string addr = (runHyp.Address ?? "").Trim();
                                                                if (string.Equals(addr, ExpectedHyperlinkAddressTask9_6, StringComparison.OrdinalIgnoreCase))
                                                                    return true;
                                                            }
                                                        }
                                                        finally
                                                        {
                                                            if (runHyp != null) { try { Marshal.ReleaseComObject(runHyp); } catch { } }
                                                        }
                                                    }
                                                }
                                                finally
                                                {
                                                    if (runAct != null) { try { Marshal.ReleaseComObject(runAct); } catch { } }
                                                }
                                            }
                                            finally
                                            {
                                                if (runActs != null) { try { Marshal.ReleaseComObject(runActs); } catch { } }
                                            }
                                        }
                                        finally
                                        {
                                            if (run != null) { try { Marshal.ReleaseComObject(run); } catch { } }
                                        }
                                        r++;
                                    }
                                }
                                catch { }
                                return false;
                            }
                            finally
                            {
                                if (tr != null) { try { Marshal.ReleaseComObject(tr); } catch { } }
                            }
                        }
                        finally
                        {
                            if (shape != null) { try { Marshal.ReleaseComObject(shape); } catch { } }
                        }
                    }
                    catch { return false; }
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>9-7: スライドの大きさを縦25.04、横33.12に変更し、スライドは画面に合わせる。</summary>
        public bool CheckTask_1_9_07()
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
                        // PowerPointの内部単位はポイント(pt)のため、cmから変換して比較します (1cm = 72/2.54 = 約28.346pt)
                        const float cmToPt = 72 / 2.54f;
                        float w = (float)pageSetup.SlideWidth;
                        float h = (float)pageSetup.SlideHeight;
                        const float expectedWidthPt = 33.12f * cmToPt;
                        const float expectedHeightPt = 25.04f * cmToPt;
                        const float tolerance = 1.0f; // 約0.35mm程度の許容範囲
                        return Math.Abs(w - expectedWidthPt) < tolerance && Math.Abs(h - expectedHeightPt) < tolerance;
                    }
                    catch { return false; }
                }
                finally { if (pageSetup != null) { try { Marshal.ReleaseComObject(pageSetup); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
    }
}
