using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Libraries;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using PptGroupShapes = Microsoft.Office.Interop.PowerPoint.GroupShapes;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題セット1 プロジェクト5（港の室内楽の会）。最終状態のみで判定する。
    /// ActivePresentation は ReleaseComObject しない（その場採点でプレゼンが閉じるのを防ぐ）。
    /// </summary>
    public class PowerPointChecker1_5_PV1
    {
        private const string SlideTitleZoom1 = "第1曲の途中では、席を探しません";
        private const string SlideTitleZoom2 = "遅れた人は、西ロビーで待ちます";
        private const string SlideTitleDecorative = "終演後のホールは、写真に残します";
        private const string SlideTitleGrayscale = "客席は、西の入口から見えます";
        private const string ExpectedFooterText = "minato-ongaku.example";
        private const int MsoSlideZoom = 36;
        private const float BelowTextTolerance = 5f;
        private const int ExpectedInverseBwMode = (int)MsoBlackWhiteMode.msoBlackWhiteInverseGrayScale;

        private static readonly Regex InverseBwModeRegex = new Regex(
            @"\bbwMode\s*=\s*""(?:invGray|inverseGray)""",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// 類1-5-1: 表紙にスライドズーム2件（指定2スライドへ）。文字の下・非重複・表紙自身を含めない。
        /// </summary>
        public bool CheckTask_1_5_01()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;

                    Slide cover = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                    if (cover == null) return false;
                    try
                    {
                        if (!TryValidateZoomPlacementOnCover(pres, cover, requiredCount: 2))
                            return false;

                        return ValidateCoverSlideZoomOpenXml(
                            pres,
                            presentationSlideNumber1Based: 1,
                            SlideTitleZoom1,
                            SlideTitleZoom2);
                    }
                    finally
                    {
                        try { Marshal.ReleaseComObject(cover); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        /// <summary>類1-5-2: 写真の代替テキストを装飾用にし、読み上げさせない。</summary>
        public bool CheckTask_1_5_02()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleDecorative);
                    if (slide == null) return false;
                    var pictures = new List<PptShape>();
                    try
                    {
                        CollectPictureCandidatesOnSlide(slide, pictures);
                        PptShape picture = GetLargestPicture(pictures);
                        if (picture == null) return false;
                        return IsShapeAltTextDecorative(picture);
                    }
                    finally
                    {
                        foreach (PptShape sh in pictures)
                        {
                            if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                        }
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        /// <summary>類1-5-3: 常に読み取り専用で開く（pptx readOnlyRecommended）。</summary>
        public bool CheckTask_1_5_03()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                Presentation pres = null;
                string tempPath = null;
                try
                {
                    pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;

                    tempPath = Path.Combine(Path.GetTempPath(), "mos_v1_5_3_" + Guid.NewGuid().ToString("N") + ".pptx");
                    try { pres.SaveCopyAs(tempPath); }
                    catch { return false; }

                    try
                    {
                        using (var package = Package.Open(tempPath, FileMode.Open, FileAccess.Read))
                        {
                            foreach (var part in package.GetParts())
                            {
                                if (!part.Uri.OriginalString.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                                    continue;
                                string xml;
                                using (var reader = new StreamReader(part.GetStream()))
                                    xml = reader.ReadToEnd();
                                if (Regex.IsMatch(xml, @"readOnlyRecommended[^>]*val\s*=\s*""(?:1|true)""", RegexOptions.IgnoreCase))
                                    return true;
                            }
                        }
                    }
                    catch { }

                    try
                    {
                        if (pres.ReadOnly == MsoTriState.msoTrue)
                            return true;
                    }
                    catch { }

                    return false;
                }
                catch { return false; }
                finally
                {
                    if (tempPath != null && File.Exists(tempPath))
                    {
                        try { File.Delete(tempPath); } catch { }
                    }
                }
            }
        }

        /// <summary>
        /// 類1-5-4: 表紙以外のフッターに「minato-ongaku.example」とスライド番号。
        /// </summary>
        public bool CheckTask_1_5_04()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slides slides = pres.Slides;
                    if (slides == null || slides.Count < 2) return false;

                    bool skipOnTitle = false;
                    try
                    {
                        if (pres.SlideMaster.HeadersFooters.DisplayOnTitleSlide == MsoTriState.msoFalse)
                            skipOnTitle = true;
                    }
                    catch { }

                    Slide slide1 = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                    try
                    {
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
                                        if (hf1.DisplayOnTitleSlide == MsoTriState.msoFalse)
                                            skipOnTitle = true;
                                    }
                                    catch { }
                                    bool footerVis = false;
                                    try { footerVis = (hf1.Footer.Visible == MsoTriState.msoTrue); } catch { }
                                    if (!footerVis) skipOnTitle = true;
                                }
                            }
                            finally
                            {
                                if (hf1 != null) { try { Marshal.ReleaseComObject(hf1); } catch { } }
                            }
                        }
                    }
                    finally
                    {
                        if (slide1 != null) { try { Marshal.ReleaseComObject(slide1); } catch { } }
                    }

                    if (!skipOnTitle) return false;

                    int last = slides.Count;
                    for (int slideNum = 2; slideNum <= last; slideNum++)
                    {
                        Slide slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
                        try
                        {
                            if (slide == null || !SlideHasFooterAndNumber(slide, ExpectedFooterText))
                                return false;
                        }
                        finally
                        {
                            if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                        }
                    }
                    return true;
                }
                catch { return false; }
            }
        }

        /// <summary>
        /// 類1-5-5: 写真を「反転させたグレースケール」に設定（最終状態 COM / OpenXML）。
        /// 表示モード切替そのものはファイルに残らないため、図形の BlackWhiteMode で判定する。
        /// </summary>
        public bool CheckTask_1_5_05()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    if (TryInverseGrayscaleViaCom())
                        return true;
                    return TryInverseGrayscaleViaOpenXml();
                }
                catch { return false; }
            }
        }

        /// <summary>
        /// 類1-5-6: 表紙音声をスライド間再生、フェードイン4秒。
        /// COM が再オープン後に取れない環境向けに VSTO 証跡（TaskV1-5-6）も見る。
        /// </summary>
        public bool CheckTask_1_5_06()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres != null)
                    {
                        bool batch = PPLogReader.IsBatchScoring();
                        int maxAttempts = batch ? 2 : 5;
                        for (int attempt = 1; attempt <= maxAttempts; attempt++)
                        {
                            Slide slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                            try
                            {
                                if (slide != null && ScanCoverAudioSettings(slide, pres))
                                    return true;
                            }
                            finally
                            {
                                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                            }

                            if (attempt < maxAttempts)
                                System.Threading.Thread.Sleep(batch ? 100 : 400);
                        }
                    }
                }
                catch { }
            }

            return PPLogReader.HasTaskV1_5_6PlayAcrossSlides()
                && PPLogReader.HasTaskV1_5_6FadeIn4000();
        }

        private static bool TryValidateZoomPlacementOnCover(Presentation pres, Slide slide, int requiredCount)
        {
            PptShapes shapes = null;
            var heldZooms = new List<PptShape>();
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
                        int visible = -999;
                        try { visible = (int)sh.Visible; } catch { }

                        // 文字の下端は「画面上に見えている文字」だけを基準にする。
                        // タイトルPH昇格で非表示の Title/Subtitle PH に文字が入っていると、
                        // 見た目より下端が深くなりズーム配置判定が落ちるため除外する。
                        if (st != MsoSlideZoom
                            && visible == (int)MsoTriState.msoTrue
                            && sh.HasTextFrame == MsoTriState.msoTrue)
                        {
                            try
                            {
                                var tf = sh.TextFrame;
                                if (tf.HasText == MsoTriState.msoTrue
                                    && tf.TextRange != null
                                    && !string.IsNullOrWhiteSpace(tf.TextRange.Text))
                                {
                                    bool isSecondaryArea = top >= slideHeight * 0.9f;
                                    bool isValidTextContainer = st == 14 || st == 17 || st == 1;
                                    if (!isSecondaryArea && isValidTextContainer)
                                    {
                                        bool isFooter = false;
                                        if (st == (int)MsoShapeType.msoPlaceholder)
                                        {
                                            try
                                            {
                                                var pf = sh.PlaceholderFormat;
                                                if (pf != null)
                                                {
                                                    try
                                                    {
                                                        PpPlaceholderType ppt = (PpPlaceholderType)pf.Type;
                                                        if (ppt == PpPlaceholderType.ppPlaceholderFooter
                                                            || ppt == PpPlaceholderType.ppPlaceholderDate
                                                            || ppt == PpPlaceholderType.ppPlaceholderSlideNumber)
                                                            isFooter = true;
                                                    }
                                                    finally { try { Marshal.ReleaseComObject(pf); } catch { } }
                                                }
                                            }
                                            catch { }
                                        }
                                        if (!isFooter)
                                        {
                                            double bottom = top + height;
                                            if (bottom > textBottom)
                                                textBottom = bottom;
                                        }
                                    }
                                }
                            }
                            catch { }
                        }

                        bool isZoomCandidate = st == MsoSlideZoom;
                        if (!isZoomCandidate)
                        {
                            try
                            {
                                string name = sh.Name ?? "";
                                string alt = sh.AlternativeText ?? "";
                                if (name.IndexOf("Zoom", StringComparison.OrdinalIgnoreCase) >= 0
                                    || name.IndexOf("ズーム", StringComparison.OrdinalIgnoreCase) >= 0
                                    || alt.IndexOf("Zoom", StringComparison.OrdinalIgnoreCase) >= 0
                                    || alt.IndexOf("ズーム", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    isZoomCandidate = true;
                                }
                            }
                            catch { }
                        }

                        if (isZoomCandidate)
                        {
                            candidateZooms.Add(Tuple.Create(sh, left, top, width, height));
                            heldZooms.Add(sh);
                            sh = null;
                        }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }

                float minTopForZoom = (float)textBottom + BelowTextTolerance;
                var zoomsBelowText = new List<Tuple<PptShape, float, float, float, float>>();
                foreach (var t in candidateZooms)
                {
                    if (t.Item3 >= minTopForZoom)
                        zoomsBelowText.Add(t);
                }

                if (zoomsBelowText.Count != requiredCount)
                    return false;

                for (int a = 0; a < zoomsBelowText.Count; a++)
                {
                    for (int b = a + 1; b < zoomsBelowText.Count; b++)
                    {
                        var ta = zoomsBelowText[a];
                        var tb = zoomsBelowText[b];
                        float la = ta.Item2, ra = ta.Item2 + ta.Item4, taTop = ta.Item3, ba = ta.Item3 + ta.Item5;
                        float lb = tb.Item2, rb = tb.Item2 + tb.Item4, tbTop = tb.Item3, bb = tb.Item3 + tb.Item5;
                        bool overlaps = !(ra <= lb || la >= rb || ba <= tbTop || taTop >= bb);
                        if (overlaps)
                            return false;
                    }
                }
                return true;
            }
            catch { return false; }
            finally
            {
                foreach (PptShape z in heldZooms)
                {
                    if (z != null) { try { Marshal.ReleaseComObject(z); } catch { } }
                }
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static bool ValidateCoverSlideZoomOpenXml(
            Presentation pres,
            int presentationSlideNumber1Based,
            string title1,
            string title2)
        {
            string tempPptxPath = null;
            try
            {
                string validationPath = ResolveValidationPptxPath(pres, "Mosapp_v1_5_1_", out tempPptxPath);
                if (string.IsNullOrWhiteSpace(validationPath))
                    return false;
                return PptxSlideZoomLinkReader.TryValidateSlideSlideZoomTargetTitles(
                    validationPath,
                    presentationSlideNumber1Based,
                    title1,
                    title2,
                    out _);
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(tempPptxPath) && File.Exists(tempPptxPath))
                {
                    try { File.Delete(tempPptxPath); } catch { }
                }
            }
        }

        private static string ResolveValidationPptxPath(Presentation pres, string tempPrefix, out string tempPptxPath)
        {
            tempPptxPath = null;
            string originalPptxPath = null;
            try { originalPptxPath = pres.FullName; } catch { }

            tempPptxPath = Path.Combine(Path.GetTempPath(), tempPrefix + Guid.NewGuid().ToString("N") + ".pptx");
            try
            {
                pres.SaveCopyAs(
                    tempPptxPath,
                    PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    MsoTriState.msoFalse);
                if (File.Exists(tempPptxPath))
                    return tempPptxPath;
            }
            catch { }

            if (string.IsNullOrWhiteSpace(originalPptxPath) || !File.Exists(originalPptxPath))
                return null;
            if (!originalPptxPath.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                return null;
            tempPptxPath = null;
            return originalPptxPath;
        }

        private static bool SlideHasFooterAndNumber(Slide slide, string expectedFooterText)
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

                bool snVis = false;
                try { snVis = (hf.SlideNumber.Visible == MsoTriState.msoTrue); } catch { }
                if (!snVis) return false;

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

        private static bool TryInverseGrayscaleViaCom()
        {
            Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
            if (pres == null) return false;
            Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleGrayscale);
            if (slide == null) return false;
            try
            {
                PptShapes shapes = slide.Shapes;
                if (shapes == null) return false;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        if (!IsPictureCandidate(sh)) continue;
                        if ((int)sh.BlackWhiteMode == ExpectedInverseBwMode)
                            return true;
                    }
                    catch { }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return false;
            }
            finally
            {
                try { Marshal.ReleaseComObject(slide); } catch { }
            }
        }

        private static bool TryInverseGrayscaleViaOpenXml()
        {
            Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
            if (pres == null) return false;

            int slideIndex;
            Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleGrayscale);
            if (slide == null) return false;
            try { slideIndex = slide.SlideIndex; }
            finally { try { Marshal.ReleaseComObject(slide); } catch { } }
            if (slideIndex < 1) return false;

            string tempPath = null;
            try
            {
                tempPath = Path.Combine(Path.GetTempPath(), "mos_v1_5_5_" + Guid.NewGuid().ToString("N") + ".pptx");
                pres.SaveCopyAs(tempPath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
                if (!File.Exists(tempPath)) return false;

                using (var package = Package.Open(tempPath, FileMode.Open, FileAccess.Read))
                {
                    string suffix = "/slides/slide" + slideIndex + ".xml";
                    var slidePart = package.GetParts()
                        .FirstOrDefault(p => p.Uri.OriginalString.EndsWith(suffix, StringComparison.OrdinalIgnoreCase));
                    if (slidePart == null) return false;

                    string slideXml;
                    using (var reader = new StreamReader(slidePart.GetStream()))
                        slideXml = reader.ReadToEnd();
                    if (string.IsNullOrEmpty(slideXml)) return false;
                    return InverseBwModeRegex.IsMatch(slideXml);
                }
            }
            catch { return false; }
            finally
            {
                if (tempPath != null && File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
        }

        private static bool ScanCoverAudioSettings(Slide slide, Presentation pres)
        {
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return false;
                int limit = 0;
                try { limit = shapes.Count; } catch { return false; }
                for (int i = 1; i <= limit; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        try { sh = shapes[i]; }
                        catch { continue; }
                        if (!PptAudioMediaHelper.TryIsSoundShape(sh))
                            continue;
                        if (!PptAudioMediaHelper.ShapeMatchesV1_5_6AudioSettings(sh, pres))
                            continue;
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

        private static void CollectPictureCandidatesOnSlide(Slide slide, List<PptShape> pictures)
        {
            if (slide == null || pictures == null) return;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        if (sh == null) continue;
                        if (sh.Type == MsoShapeType.msoGroup)
                        {
                            PptGroupShapes group = null;
                            try
                            {
                                group = sh.GroupItems;
                                if (group == null) continue;
                                for (int g = 1; g <= group.Count; g++)
                                {
                                    PptShape member = null;
                                    try
                                    {
                                        member = group[g];
                                        if (IsPictureCandidate(member))
                                        {
                                            pictures.Add(member);
                                            member = null;
                                        }
                                    }
                                    finally
                                    {
                                        if (member != null) { try { Marshal.ReleaseComObject(member); } catch { } }
                                    }
                                }
                            }
                            finally
                            {
                                if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
                            }
                            continue;
                        }

                        if (IsPictureCandidate(sh))
                        {
                            pictures.Add(sh);
                            sh = null;
                        }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
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
                        if (pf.ContainedType == MsoShapeType.msoPicture
                            || pf.Type == PpPlaceholderType.ppPlaceholderPicture)
                            return true;
                    }
                    finally { try { Marshal.ReleaseComObject(pf); } catch { } }
                }
            }
            catch { }

            try
            {
                string name = sh.Name ?? "";
                name = name.ToLowerInvariant();
                return name.Contains("picture") || name.Contains("画像") || name.Contains("図");
            }
            catch { return false; }
        }

        private static PptShape GetLargestPicture(List<PptShape> pictures)
        {
            if (pictures == null || pictures.Count == 0) return null;
            PptShape largest = pictures[0];
            float maxArea = GetShapeArea(largest);
            for (int i = 1; i < pictures.Count; i++)
            {
                float area = GetShapeArea(pictures[i]);
                if (area > maxArea)
                {
                    maxArea = area;
                    largest = pictures[i];
                }
            }
            return largest;
        }

        private static float GetShapeArea(PptShape sh)
        {
            if (sh == null) return 0f;
            try { return (float)sh.Width * (float)sh.Height; }
            catch { return 0f; }
        }

        private static bool IsShapeAltTextDecorative(PptShape sh)
        {
            if (sh == null) return false;

            string alt = "";
            string title = "";
            try { alt = sh.AlternativeText ?? ""; } catch { }
            try { title = sh.Title ?? ""; } catch { }

            try
            {
                dynamic dSh = sh;
                int decorativeValue = (int)dSh.Decorative;
                if (decorativeValue == -1) return true;
            }
            catch { }

            return string.IsNullOrWhiteSpace(alt)
                && title.IndexOf("Decorative", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
