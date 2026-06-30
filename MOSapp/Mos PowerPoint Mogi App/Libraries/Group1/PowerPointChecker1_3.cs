using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using PptThreeDFormat = Microsoft.Office.Interop.PowerPoint.ThreeDFormat;

namespace Libraries.Group1
{
    public class PowerPointChecker1_3
    {
        private const int MsoSlideZoom = 36;
        private const int MsoSectionZoom = 37;
        private const float BelowTextTolerance = 5f;

        /// <summary>P3-1: スライド7にSmartArt「タイムライン」で「コンテンツ」「デザイン」を入力。</summary>
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
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 7);
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
                            string layoutId = layout.Id ?? "";
                            string layoutName = "";
                            try { layoutName = layout.Name ?? ""; } catch { }
                            if (!IsTimelineSmartArtLayout(layoutId, layoutName))
                                return false;
                        }
                        finally
                        {
                            if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } }
                        }

                        string allText = CollectSmartArtText(smartArt);
                        return allText.IndexOf("コンテンツ", StringComparison.OrdinalIgnoreCase) >= 0
                            && allText.IndexOf("デザイン", StringComparison.OrdinalIgnoreCase) >= 0;
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
        /// 「タイムライン」系 SmartArt レイアウトか。基本タイムラインは Id が hProcess11 で Timeline 文字列を含まない。
        /// </summary>
        private static bool IsTimelineSmartArtLayout(string layoutId, string layoutName)
        {
            if (!string.IsNullOrEmpty(layoutId))
            {
                if (layoutId.IndexOf("Timeline", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                if (layoutId.IndexOf("hProcess11", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            if (!string.IsNullOrEmpty(layoutName))
            {
                if (layoutName.IndexOf("タイムライン", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                if (layoutName.IndexOf("Timeline", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
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
                        AppendSmartArtNodeText(node, sb);
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

        private static void AppendSmartArtNodeText(SmartArtNode node, StringBuilder sb)
        {
            try
            {
                var tf2 = node.TextFrame2;
                if (tf2 != null && tf2.TextRange != null)
                    sb.Append(tf2.TextRange.Text ?? "");
            }
            catch { }
        }

        /// <summary>P3-2: スライド7のSmartArtの色を「グラデーション循環-アクセント6」に変更。</summary>
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
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 7);
                    if (slide == null) return false;
                    saShape = PowerPointCheckerCommon.FindSmartArtShape(slide);
                    if (saShape == null) return false;
                    SmartArt smartArt = null;
                    try
                    {
                        smartArt = saShape.SmartArt;
                        if (smartArt == null) return false;
                        SmartArtColor appliedColor = null;
                        try
                        {
                            appliedColor = smartArt.Color;
                            if (appliedColor == null) return false;
                            string appliedId = appliedColor.Id ?? "";
                            string appliedName = appliedColor.Name ?? "";

                            if (IsP3_2GradientCycleAccent6(appliedId, appliedName))
                                return true;

                            for (int idx = 1; idx <= 20; idx++)
                            {
                                try
                                {
                                    var style = app.SmartArtColors[idx];
                                    if (style == null) continue;
                                    string styleId = style.Id ?? "";
                                    string styleName = style.Name ?? "";
                                    if (IsP3_2GradientCycleAccent6(styleId, styleName) && styleId == appliedId)
                                    {
                                        try { Marshal.ReleaseComObject(style); } catch { }
                                        return true;
                                    }
                                    try { Marshal.ReleaseComObject(style); } catch { }
                                }
                                catch { break; }
                            }
                            return false;
                        }
                        finally
                        {
                            if (appliedColor != null) { try { Marshal.ReleaseComObject(appliedColor); } catch { } }
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

        /// <summary>
        /// P3-2 正答色「グラデーション 循環 アクセント6」のみ許容。
        /// 「グラデーション アクセント6」「グラデーション 透過 アクセント6」は除外する。
        /// </summary>
        private static bool IsP3_2GradientCycleAccent6(string id, string name)
        {
            string combined = ((id ?? "") + " " + (name ?? "")).ToLowerInvariant();
            bool accent6 = combined.IndexOf("accent6", StringComparison.Ordinal) >= 0
                || combined.IndexOf("accent 6", StringComparison.Ordinal) >= 0
                || combined.IndexOf("アクセント6", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("アクセント 6", StringComparison.OrdinalIgnoreCase) >= 0
                || combined.IndexOf("アクセント６", StringComparison.OrdinalIgnoreCase) >= 0;
            bool gradient = combined.IndexOf("gradient", StringComparison.Ordinal) >= 0
                || combined.IndexOf("グラデーション", StringComparison.OrdinalIgnoreCase) >= 0;
            bool cycle = combined.IndexOf("cycle", StringComparison.Ordinal) >= 0
                || combined.IndexOf("循環", StringComparison.OrdinalIgnoreCase) >= 0;
            bool transparent = combined.IndexOf("transparent", StringComparison.Ordinal) >= 0
                || combined.IndexOf("透過", StringComparison.OrdinalIgnoreCase) >= 0;
            return accent6 && gradient && cycle && !transparent;
        }

        /// <summary>P3-3: スライド6の箇条書きを「ターゲットリスト」のSmartArtに変更。</summary>
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
                        SmartArtLayout layout = null;
                        try
                        {
                            layout = smartArt.Layout;
                            if (layout == null) return false;
                            string layoutId = layout.Id ?? "";
                            string layoutName = "";
                            try { layoutName = layout.Name ?? ""; } catch { }
                            return IsTargetListSmartArtLayout(layoutId, layoutName);
                        }
                        finally
                        {
                            if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } }
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

        /// <summary>
        /// 「ターゲットリスト」系 SmartArt レイアウトか。Id は target3 で TargetList 文字列を含まない。
        /// </summary>
        private static bool IsTargetListSmartArtLayout(string layoutId, string layoutName)
        {
            if (!string.IsNullOrEmpty(layoutId))
            {
                if (layoutId.IndexOf("TargetList", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                if (layoutId.IndexOf("/target3", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            if (!string.IsNullOrEmpty(layoutName))
            {
                if (layoutName.IndexOf("ターゲットリスト", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                if (layoutName.IndexOf("Target List", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            return false;
        }

        /// <summary>P3-4: スライド1に3Dモデル「虫眼鏡」を挿入し、幅2.5cmに変更。</summary>
        public bool CheckTask_1_3_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                PptShape modelShape = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                    if (slide == null) return false;

                    modelShape = PowerPointCheckerCommon.Find3DModelShapeByName(slide, "虫眼鏡");
                    if (modelShape == null)
                        modelShape = PowerPointCheckerCommon.Find3DModelShapeByName(slide, "Magnifying");
                    if (modelShape == null)
                        modelShape = PowerPointCheckerCommon.Find3DModelShape(slide);
                    if (modelShape == null) return false;

                    float widthCm = (float)modelShape.Width * 2.54f / 72f;
                    return Math.Abs(widthCm - 2.5f) <= 0.1f;
                }
                finally
                {
                    if (modelShape != null) { try { Marshal.ReleaseComObject(modelShape); } catch { } }
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P3-5: スライド10の3Dモデルのビューを上前面にし、高さ6.5cmに変更。</summary>
        public bool CheckTask_1_3_05()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                PptShape modelShape = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 10);
                    if (slide == null) return false;
                    modelShape = PowerPointCheckerCommon.Find3DModelShape(slide);
                    if (modelShape == null) return false;

                    if (!IsTopFront3DView(modelShape))
                        return false;

                    float heightCm = (float)modelShape.Height * 2.54f / 72f;
                    return Math.Abs(heightCm - 6.5f) <= 0.15f;
                }
                finally
                {
                    if (modelShape != null) { try { Marshal.ReleaseComObject(modelShape); } catch { } }
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P3-6: 「機能の概要」スライドに3件のスライドズームを挿入。</summary>
        public bool CheckTask_1_3_06()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitlePlaceholderExact(pres, "機能の概要");
                    if (slide == null) return false;

                    int slideNumber = slide.SlideIndex;
                    var requiredTitles = new[]
                    {
                        "画面録画で説明",
                        "ズーム機能で訴求力アップ",
                        "デザインアイデアで魅力的に！"
                    };

                    if (!TryValidateZoomPlacementOnSlide(
                            pres, slide, MsoSlideZoom, requiredTitles.Length,
                            titlePlaceholderOnlyForTextBaseline: true,
                            out _, out _))
                        return false;

                    return ValidateSlideZoomOpenXml(pres, slideNumber, requiredTitles, out _);
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P3-7: スライド2にセクションズームを2件挿入。</summary>
        public bool CheckTask_1_3_07()
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

                    return ValidateSectionZoomPlacedUnderLabels(
                        pres, 2,
                        "1.機能の概要", "1.機能の概要",
                        "2.伝わるスライドの要素", "2.伝わるスライドの要素",
                        out _);
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        private static bool IsTopFront3DView(PptShape modelShape)
        {
            PptThreeDFormat threeD = null;
            try
            {
                threeD = modelShape.ThreeD;
                if (threeD == null) return false;
                double rotX = (double)threeD.RotationX;
                double rotY = (double)threeD.RotationY;
                // 旧6-4（下背面）: rotX≈20, rotY≈180。上前面は Y が 180° 付近でない。
                if (Math.Abs(rotY - 180.0) < 20.0) return false;
                if (rotX < -40.0 || rotX > 15.0) return false;
                return true;
            }
            catch { return false; }
            finally
            {
                if (threeD != null) { try { Marshal.ReleaseComObject(threeD); } catch { } }
            }
        }

        private static bool TryValidateZoomPlacementOnSlide(
            Presentation pres,
            Slide slide,
            int zoomShapeType,
            int requiredCount,
            bool titlePlaceholderOnlyForTextBaseline,
            out string failReason,
            out List<Tuple<PptShape, float, float, float, float>> zoomsBelowText)
        {
            failReason = null;
            zoomsBelowText = null;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null)
                {
                    failReason = "Shapes is null";
                    return false;
                }

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

                        if (st != zoomShapeType && sh.HasTextFrame == MsoTriState.msoTrue)
                        {
                            try
                            {
                                var tf = sh.TextFrame;
                                if (tf.HasText == MsoTriState.msoTrue && tf.TextRange != null && !string.IsNullOrWhiteSpace(tf.TextRange.Text))
                                {
                                    bool isSecondaryArea = top >= slideHeight * 0.9f;
                                    bool isValidTextContainer = st == 14 || st == 17 || st == 1;
                                    if (!isSecondaryArea && isValidTextContainer)
                                    {
                                        bool isFooter = false;
                                        bool isBodyPlaceholder = false;
                                        if (st == (int)MsoShapeType.msoPlaceholder)
                                        {
                                            try
                                            {
                                                var pf = sh.PlaceholderFormat;
                                                if (pf != null)
                                                {
                                                    PpPlaceholderType ppt = (PpPlaceholderType)pf.Type;
                                                    if (ppt == PpPlaceholderType.ppPlaceholderFooter
                                                        || ppt == PpPlaceholderType.ppPlaceholderDate
                                                        || ppt == PpPlaceholderType.ppPlaceholderSlideNumber)
                                                        isFooter = true;
                                                    if (titlePlaceholderOnlyForTextBaseline
                                                        && !IsTitleLikePlaceholder(ppt))
                                                        isBodyPlaceholder = true;
                                                    try { Marshal.ReleaseComObject(pf); } catch { }
                                                }
                                            }
                                            catch { }
                                        }
                                        if (!isFooter && !isBodyPlaceholder)
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

                        bool isZoomCandidate = st == zoomShapeType;
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
                            sh = null;
                        }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }

                float belowTolerance = titlePlaceholderOnlyForTextBaseline ? 0f : BelowTextTolerance;
                float minTopForZoom = (float)textBottom + belowTolerance;
                zoomsBelowText = new List<Tuple<PptShape, float, float, float, float>>();
                foreach (var t in candidateZooms)
                {
                    if (t.Item3 >= minTopForZoom)
                        zoomsBelowText.Add(t);
                    else
                    {
                        try { Marshal.ReleaseComObject(t.Item1); } catch { }
                    }
                }

                if (zoomsBelowText.Count != requiredCount)
                {
                    failReason = $"zoom count below title baseline expected={requiredCount} candidates={candidateZooms.Count} belowBaseline={zoomsBelowText.Count} textBottom={textBottom:F1} minTop={minTopForZoom:F1} belowTolerance={belowTolerance:F1} titleOnlyBaseline={titlePlaceholderOnlyForTextBaseline}";
                    ReleaseZoomShapes(zoomsBelowText);
                    zoomsBelowText = null;
                    return false;
                }

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
                        {
                            failReason = "zoom shapes overlap";
                            ReleaseZoomShapes(zoomsBelowText);
                            zoomsBelowText = null;
                            return false;
                        }
                    }
                }

                ReleaseZoomShapes(zoomsBelowText);
                zoomsBelowText = null;
                return true;
            }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static bool IsTitleLikePlaceholder(PpPlaceholderType placeholderType)
        {
            return placeholderType == PpPlaceholderType.ppPlaceholderTitle
                || placeholderType == PpPlaceholderType.ppPlaceholderCenterTitle
                || placeholderType == PpPlaceholderType.ppPlaceholderVerticalTitle
                || placeholderType == PpPlaceholderType.ppPlaceholderSubtitle;
        }

        private static void ReleaseZoomShapes(List<Tuple<PptShape, float, float, float, float>> zooms)
        {
            if (zooms == null) return;
            foreach (var t in zooms)
            {
                try { Marshal.ReleaseComObject(t.Item1); } catch { }
            }
        }

        private static bool ValidateSlideZoomOpenXml(Presentation pres, int slideNumber, IReadOnlyList<string> requiredTitles)
        {
            return ValidateSlideZoomOpenXml(pres, slideNumber, requiredTitles, out _);
        }

        private static bool ValidateSlideZoomOpenXml(
            Presentation pres,
            int slideNumber,
            IReadOnlyList<string> requiredTitles,
            out string errorMessage)
        {
            errorMessage = null;
            string tempPptxPath = null;
            try
            {
                string validationPptxPath = ResolveValidationPptxPath(pres, "Mosapp_3_6_", out tempPptxPath);
                if (string.IsNullOrWhiteSpace(validationPptxPath))
                {
                    errorMessage = "validation pptx path unavailable";
                    return false;
                }

                return PptxSlideZoomLinkReader.TryValidateSlideSlideZoomTargetTitles(
                    validationPptxPath,
                    slideNumber,
                    requiredTitles,
                    out errorMessage);
            }
            finally
            {
                DeleteTempPptx(tempPptxPath);
            }
        }

        private static bool ValidateSectionZoomPlacedUnderLabels(
            Presentation pres,
            int slideNumber,
            string labelText1,
            string expectedSectionName1,
            string labelText2,
            string expectedSectionName2,
            out string errorMessage)
        {
            errorMessage = null;
            string tempPptxPath = null;
            try
            {
                string validationPptxPath = ResolveValidationPptxPath(pres, "Mosapp_3_7_", out tempPptxPath);
                if (string.IsNullOrWhiteSpace(validationPptxPath))
                {
                    errorMessage = "validation pptx path unavailable";
                    return false;
                }

                return PptxSlideZoomLinkReader.TryValidateSectionZoomPlacedUnderLabels(
                    validationPptxPath,
                    slideNumber,
                    labelText1,
                    expectedSectionName1,
                    labelText2,
                    expectedSectionName2,
                    out errorMessage);
            }
            finally
            {
                DeleteTempPptx(tempPptxPath);
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

        private static void DeleteTempPptx(string tempPptxPath)
        {
            if (string.IsNullOrWhiteSpace(tempPptxPath)) return;
            try { if (File.Exists(tempPptxPath)) File.Delete(tempPptxPath); } catch { }
        }
    }
}
