using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Libraries;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題セット1 プロジェクト3（釜辺の手元）。最終状態のみで判定する。
    /// ActivePresentation / Designs / DocumentProperties / SmartArtColors は
    /// ReleaseComObject するとその場採点でプレゼンが閉じることがあるため解放しない。
    /// </summary>
    public class PowerPointChecker1_3_PV1
    {
        private const string SlideTitleTable = "陶芸は16名、書道は28名です";
        private const string SlideTitleSmartArt = "どちらも、3回で1作目が残ります";
        private const string SlideTitlePhotos = "左がロクロ、右が半紙です";
        private const string SlideTitleLink = "空きは、ページで確認します";
        private const string ExpectedHyperlinkText = "申し込み";
        private const string ExpectedHyperlinkAddress = "https://www.kamabe-temoto.example/";
        private const string ExpectedThemeNameJa = "イオン";
        private const string ExpectedThemeNameEn = "Ion";
        private const float PositionTolerance = 2.0f;

        /// <summary>
        /// 類1-3-1: スライド「陶芸は16名、書道は28名です」のレイアウトを「表スライド」に変更。
        /// （新規スライド追加ではなく、既存スライドのレイアウト変更）
        /// </summary>
        public bool CheckTask_1_3_01()
        {
            // テーマ（1-3-6）適用後は CustomLayout 名「表スライド」が消えることがある。
            // 類題専用 VSTO 証跡 [TaskV1-3-1] TableSlideLayout があれば合格とする。
            if (PPLogReader.HasTaskV1_3_1TableSlideLayout())
                return true;

            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleTable);
                    if (slide == null) return false;
                    try
                    {
                        CustomLayout layout = slide.CustomLayout;
                        if (layout == null) return false;
                        string layoutName = "";
                        try { layoutName = layout.Name ?? ""; } catch { return false; }
                        return layoutName.IndexOf("表スライド", StringComparison.OrdinalIgnoreCase) >= 0
                            || layoutName.IndexOf("テーブルスライド", StringComparison.OrdinalIgnoreCase) >= 0;
                    }
                    finally
                    {
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        public bool CheckTask_1_3_02()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleSmartArt);
                    if (slide == null) return false;
                    PptShape saShape = null;
                    try
                    {
                        saShape = PowerPointCheckerCommon.FindSmartArtShape(slide);
                        if (saShape == null) return false;
                        SmartArt smartArt = saShape.SmartArt;
                        if (smartArt == null) return false;
                        SmartArtColor appliedColor = smartArt.Color;
                        if (appliedColor == null) return false;
                        string appliedId = "";
                        string appliedName = "";
                        try { appliedId = appliedColor.Id ?? ""; } catch { }
                        try { appliedName = appliedColor.Name ?? ""; } catch { }
                        // Application.SmartArtColors の列挙＋Release はプレゼン閉鎖の原因になるため行わない。
                        return IsFillAccent5Color(appliedId, appliedName);
                    }
                    finally
                    {
                        if (saShape != null) { try { Marshal.ReleaseComObject(saShape); } catch { } }
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        public bool CheckTask_1_3_03()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitlePhotos);
                    if (slide == null) return false;
                    try
                    {
                        if (!TryGetLeftAndRightPictures(slide, out PptShape leftPic, out PptShape rightPic))
                            return false;
                        try
                        {
                            float leftTop = (float)leftPic.Top;
                            float rightTop = (float)rightPic.Top;
                            return Math.Abs(rightTop - leftTop) <= PositionTolerance;
                        }
                        finally
                        {
                            if (leftPic != null) { try { Marshal.ReleaseComObject(leftPic); } catch { } }
                            if (rightPic != null) { try { Marshal.ReleaseComObject(rightPic); } catch { } }
                        }
                    }
                    finally
                    {
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        public bool CheckTask_1_3_04()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;

                    Slides slides = pres.Slides;
                    if (slides == null) return false;
                    for (int i = 1; i <= slides.Count; i++)
                    {
                        Slide slide = slides[i];
                        try
                        {
                            Comments comments = slide.Comments;
                            if (comments != null && comments.Count > 0)
                                return false;
                        }
                        finally
                        {
                            try { Marshal.ReleaseComObject(slide); } catch { }
                        }
                    }

                    // BuiltInDocumentProperties は ReleaseComObject しない（pres 解放と合わせて閉鎖しやすい）。
                    dynamic props = pres.BuiltInDocumentProperties;
                    if (props == null) return true;
                    string[] personalPropNames =
                    {
                        "Author", "Manager", "Company", "Last Author",
                        "Title", "Subject", "Keywords", "Comments"
                    };
                    foreach (string name in personalPropNames)
                    {
                        try
                        {
                            object val = props[name].Value;
                            string s = (val == null) ? "" : (val.ToString() ?? "").Trim();
                            if (!string.IsNullOrEmpty(s))
                                return false;
                        }
                        catch { }
                    }
                    return true;
                }
                catch { return false; }
            }
        }

        public bool CheckTask_1_3_05()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleLink);
                    if (slide == null) return false;
                    PptShape shape = null;
                    try
                    {
                        shape = PowerPointCheckerCommon.FindShapeWithText(slide, ExpectedHyperlinkText);
                        if (shape == null) return false;
                        Microsoft.Office.Interop.PowerPoint.TextFrame tf = shape.TextFrame;
                        if (tf == null) return false;
                        TextRange tr = tf.TextRange;
                        if (tr == null) return false;
                        if (TryHyperlinkOnTextRange(tr, ExpectedHyperlinkAddress))
                            return true;
                        return TryHyperlinkOnRunsContainingText(tr, ExpectedHyperlinkText, ExpectedHyperlinkAddress);
                    }
                    finally
                    {
                        if (shape != null) { try { Marshal.ReleaseComObject(shape); } catch { } }
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        public bool CheckTask_1_3_06()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;

                    // Designs / Design は ReleaseComObject しない（テーマ関連 COM 解放でプレゼンが閉じることがある）。
                    Designs designs = pres.Designs;
                    if (designs != null)
                    {
                        for (int i = 1; i <= designs.Count; i++)
                        {
                            try
                            {
                                Design design = designs[i];
                                string name = "";
                                try { name = design.Name ?? ""; } catch { }
                                if (IsIonThemeName(name))
                                    return true;
                            }
                            catch { }
                        }
                    }

                    Slide slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                    if (slide == null) return false;
                    try
                    {
                        Slides slides = pres.Slides;
                        if (slides == null) return false;
                        SlideRange range = slides.Range(new object[] { slide.SlideIndex });
                        if (range == null) return false;
                        Design slideDesign = range.Design;
                        if (slideDesign == null) return false;
                        string slideDesignName = "";
                        try { slideDesignName = slideDesign.Name ?? ""; } catch { return false; }
                        return IsIonThemeName(slideDesignName);
                    }
                    finally
                    {
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        private static bool IsIonThemeName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            return name.IndexOf(ExpectedThemeNameJa, StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf(ExpectedThemeNameEn, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>「塗りつぶし - アクセント 5」(accent5_2)。枠線のみ／グラデーション／カラフルは除外。</summary>
        private static bool IsFillAccent5Color(string id, string name)
        {
            string combined = ((id ?? "") + " " + (name ?? "")).ToLowerInvariant();
            if (combined.IndexOf("accent5_2", StringComparison.Ordinal) >= 0)
                return true;

            bool accent5 = combined.IndexOf("accent5", StringComparison.Ordinal) >= 0
                || combined.IndexOf("accent 5", StringComparison.Ordinal) >= 0
                || combined.IndexOf("アクセント5", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("アクセント 5", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("アクセント５", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!accent5) return false;

            bool outlineOnly = combined.IndexOf("枠線のみ", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("outline", StringComparison.Ordinal) >= 0
                || combined.IndexOf("accent5_1", StringComparison.Ordinal) >= 0;
            if (outlineOnly) return false;

            bool gradient = combined.IndexOf("gradient", StringComparison.Ordinal) >= 0
                || combined.IndexOf("グラデーション", StringComparison.OrdinalIgnoreCase) >= 0;
            if (gradient) return false;

            bool colorful = combined.IndexOf("colorful", StringComparison.Ordinal) >= 0
                || combined.IndexOf("カラフル", StringComparison.OrdinalIgnoreCase) >= 0;
            if (colorful) return false;

            return combined.IndexOf("塗りつぶし", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("fill", StringComparison.Ordinal) >= 0;
        }

        private static bool TryGetLeftAndRightPictures(Slide slide, out PptShape leftPic, out PptShape rightPic)
        {
            leftPic = null;
            rightPic = null;
            if (slide == null) return false;

            var pictures = new List<PptShape>();
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
                        if (!IsPictureCandidate(sh)) continue;
                        pictures.Add(sh);
                        sh = null;
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }

                if (pictures.Count < 2) return false;

                int leftIdx = 0;
                int rightIdx = 0;
                float minLeft = (float)pictures[0].Left;
                float maxLeft = minLeft;
                for (int i = 1; i < pictures.Count; i++)
                {
                    float left = (float)pictures[i].Left;
                    if (left < minLeft)
                    {
                        minLeft = left;
                        leftIdx = i;
                    }
                    if (left > maxLeft)
                    {
                        maxLeft = left;
                        rightIdx = i;
                    }
                }

                if (leftIdx == rightIdx || Math.Abs(maxLeft - minLeft) < 0.5f)
                    return false;

                leftPic = pictures[leftIdx];
                rightPic = pictures[rightIdx];
                for (int i = 0; i < pictures.Count; i++)
                {
                    if (i == leftIdx || i == rightIdx) continue;
                    try { Marshal.ReleaseComObject(pictures[i]); } catch { }
                }
                return true;
            }
            catch
            {
                foreach (var sh in pictures)
                {
                    if (sh == null || ReferenceEquals(sh, leftPic) || ReferenceEquals(sh, rightPic)) continue;
                    try { Marshal.ReleaseComObject(sh); } catch { }
                }
                leftPic = null;
                rightPic = null;
                return false;
            }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static bool IsPictureCandidate(PptShape sh)
        {
            if (sh == null) return false;
            if (PowerPointCheckerCommon.IsPictureShape(sh))
                return true;
            if (sh.Type != MsoShapeType.msoPlaceholder)
                return false;

            try
            {
                PlaceholderFormat pf = sh.PlaceholderFormat;
                if (pf != null)
                {
                    try
                    {
                        if (pf.ContainedType == MsoShapeType.msoPicture)
                            return true;
                    }
                    finally { try { Marshal.ReleaseComObject(pf); } catch { } }
                }
            }
            catch { }
            return false;
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
                    string runText = "";
                    try { runText = run.Text ?? ""; } catch { }
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
            try
            {
                ActionSettings acts = tr.ActionSettings;
                if (acts == null) return false;
                ActionSetting act = acts[PpMouseActivation.ppMouseClick];
                if (act == null) return false;
                if (act.Action != PpActionType.ppActionHyperlink) return false;
                Hyperlink hyp = act.Hyperlink;
                if (hyp == null) return false;
                string addr = "";
                try { addr = hyp.Address ?? ""; } catch { }
                return string.Equals(addr.Trim(), (expectedAddress ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }
    }
}
