using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using Libraries;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_4
    {
        private const double PositionTolerance = 2.0;

        /// <summary>P4-1: スライド1の青い図形に「教育者必見」。</summary>
        public bool CheckTask_1_4_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                PptShape shape = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                    if (slide == null) return false;
                    shape = PowerPointCheckerCommon.FindShapeWithText(slide, "教育者必見");
                    return shape != null;
                }
                finally
                {
                    if (shape != null) { try { Marshal.ReleaseComObject(shape); } catch { } }
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P4-2: スライド4の「いつでも体験可能です!」の塗りつぶしを「青、アクセント1」。</summary>
        public bool CheckTask_1_4_02()
        {
            const string targetText = "いつでも体験可能です!";
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

                    PptShape shDeep = null;
                    PptShape shShallow = null;
                    try
                    {
                        shDeep = PowerPointCheckerCommon.FindShapeWithTextDeep(slide, targetText);
                        shShallow = PowerPointCheckerCommon.FindShapeWithText(slide, targetText);
                        PptShape sh = shDeep ?? shShallow;
                        if (sh == null) return false;

                        return PowerPointCheckerCommon.IsShapeSubstringFontThemeColor(
                            sh, pres, targetText, MsoThemeColorIndex.msoThemeColorAccent1);
                    }
                    finally
                    {
                        if (shShallow != null && !ReferenceEquals(shShallow, shDeep))
                        {
                            try { Marshal.ReleaseComObject(shShallow); } catch { }
                        }
                        if (shDeep != null) { try { Marshal.ReleaseComObject(shDeep); } catch { } }
                    }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P4-3: スライド1の子供の画像に光彩 18pt・緑アクセントカラー6。</summary>
        public bool CheckTask_1_4_03()
        {
            if (PPLogReader.HasTask4_3GlowExecuted())
                return true;

            Presentation pres = null;
            string tempPath = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                    if (slide == null) return false;

                    if (TryValidateGlowOnAnyPicture(slide))
                        return true;

                    tempPath = Path.Combine(Path.GetTempPath(), "mos_4_3_check_" + Guid.NewGuid().ToString("N") + ".pptx");
                    pres.SaveCopyAs(tempPath);
                    return PptxGlowOpenXmlReader.ContainsGlow18ptAccent6OnSlide(tempPath, 1);
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
                if (tempPath != null && File.Exists(tempPath))
                    try { File.Delete(tempPath); } catch { }
            }
        }

        private static bool TryValidateGlowOnAnyPicture(Slide slide)
        {
            if (slide == null) return false;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return false;
                return TryValidateGlowInShapeCollection(shapes, searchGroups: true);
            }
            catch { return false; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static bool TryValidateGlowInShapeCollection(PptShapes shapes, bool searchGroups)
        {
            if (shapes == null) return false;
            int count = shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = shapes[i];
                    if (TryValidateGlowOnPictureShape(sh))
                        return true;
                    if (searchGroups && sh.Type == MsoShapeType.msoGroup)
                    {
                        Microsoft.Office.Interop.PowerPoint.GroupShapes group = null;
                        try
                        {
                            group = sh.GroupItems;
                            if (group != null && TryValidateGlowInGroupShapes(group))
                                return true;
                        }
                        finally
                        {
                            if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
                        }
                    }
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
            return false;
        }

        private static bool TryValidateGlowInGroupShapes(Microsoft.Office.Interop.PowerPoint.GroupShapes group)
        {
            if (group == null) return false;
            int count = group.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = group[i];
                    if (TryValidateGlowOnPictureShape(sh))
                        return true;
                    if (sh.Type == MsoShapeType.msoGroup)
                    {
                        Microsoft.Office.Interop.PowerPoint.GroupShapes nested = null;
                        try
                        {
                            nested = sh.GroupItems;
                            if (nested != null && TryValidateGlowInGroupShapes(nested))
                                return true;
                        }
                        finally
                        {
                            if (nested != null) { try { Marshal.ReleaseComObject(nested); } catch { } }
                        }
                    }
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
            return false;
        }

        private static bool TryValidateGlowOnPictureShape(PptShape sh)
        {
            if (sh == null || !IsPictureCandidate(sh))
                return false;

            dynamic glow = null;
            try
            {
                glow = sh.Glow;
                if (glow == null) return false;

                float radius = 0f;
                try { radius = (float)glow.Radius; } catch { }
                if (!IsGlowRadius18pt(radius))
                    return false;

                Microsoft.Office.Interop.PowerPoint.ColorFormat cf = null;
                try
                {
                    cf = glow.Color;
                    return IsGlowAccent6Color(cf);
                }
                finally
                {
                    if (cf != null) { try { Marshal.ReleaseComObject(cf); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (glow != null) { try { Marshal.ReleaseComObject(glow); } catch { } }
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
                var pf = sh.PlaceholderFormat;
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

            try
            {
                string name = sh.Name ?? "";
                name = name.ToLowerInvariant();
                return name.Contains("picture") || name.Contains("画像") || name.Contains("図");
            }
            catch { return false; }
        }

        private static bool IsGlowRadius18pt(float radius)
        {
            if (radius <= 0f) return false;
            // UI 18pt。環境差を吸収（14〜22pt 相当）
            return radius >= 14f && radius <= 22f;
        }

        private static bool IsGlowAccent6Color(Microsoft.Office.Interop.PowerPoint.ColorFormat cf)
        {
            if (cf == null) return false;
            try
            {
                if (cf.ObjectThemeColor == MsoThemeColorIndex.msoThemeColorAccent6)
                    return true;
            }
            catch { }

            try
            {
                if (cf.Type != MsoColorType.msoColorTypeRGB)
                    return false;
                int rgb = (int)cf.RGB;
                int r = rgb & 0xFF;
                int g = (rgb >> 8) & 0xFF;
                int b = (rgb >> 16) & 0xFF;
                // 既定テーマのアクセント6（緑系）および近傍色
                if (r >= 60 && r <= 140 && g >= 140 && g <= 210 && b >= 40 && b <= 120)
                    return true;
            }
            catch { }

            return false;
        }

        /// <summary>P4-4: スライド1の画像に「楕円 ぼかし」スタイル＋「テクスチャライザー」アート効果。</summary>
        public bool CheckTask_1_4_04()
        {
            Presentation pres = null;
            string tempPath = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                bool styleOk = false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                    if (slide != null)
                    {
                        PptShape picture = GetFirstPictureOnSlide(slide);
                        if (picture != null)
                        {
                            try
                            {
                                dynamic dSh = picture;
                                try
                                {
                                    dynamic se = dSh.SoftEdge;
                                    if (se != null)
                                    {
                                        if ((int)se.Type >= 1 || (float)se.Radius > 0) styleOk = true;
                                        Marshal.ReleaseComObject(se);
                                    }
                                }
                                catch { }
                                if (!styleOk)
                                {
                                    try
                                    {
                                        dynamic pf = dSh.PictureFormat;
                                        int ps = (int)pf.PictureStyle;
                                        if (ps == 13 || ps == 20 || ps == 21 || ps == 22) styleOk = true;
                                        Marshal.ReleaseComObject(pf);
                                    }
                                    catch { }
                                }
                            }
                            finally { Marshal.ReleaseComObject(picture); }
                        }
                    }
                }
                finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }

                bool effectOk = false;
                try
                {
                    tempPath = Path.Combine(Path.GetTempPath(), "mos_4_4_check_" + Guid.NewGuid().ToString("N") + ".pptx");
                    pres.SaveCopyAs(tempPath);
                    using (var package = Package.Open(tempPath, FileMode.Open, FileAccess.Read))
                    {
                        foreach (var part in package.GetParts())
                        {
                            if (!part.Uri.OriginalString.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;
                            string xml;
                            try
                            {
                                using (var reader = new StreamReader(part.GetStream()))
                                    xml = reader.ReadToEnd();
                            }
                            catch { continue; }

                            if (Regex.IsMatch(xml, @"artisticTexturizer|Texturizer", RegexOptions.IgnoreCase))
                            {
                                effectOk = true;
                                break;
                            }
                        }
                    }
                }
                catch { }

                return styleOk && effectOk;
            }
            catch { return false; }
            finally
            {
                if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { }
                if (tempPath != null && File.Exists(tempPath))
                    try { File.Delete(tempPath); } catch { }
            }
        }

        /// <summary>P4-5: スライド5の右の画像を左の画像の上端に合わせる（水平位置は不変）。</summary>
        public bool CheckTask_1_4_05()
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
                finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }
            }
            catch { return false; }
            finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
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
                CollectPictureCandidates(shapes, pictures, searchGroups: true);
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
                if (leftPic != null) { try { Marshal.ReleaseComObject(leftPic); } catch { } leftPic = null; }
                if (rightPic != null) { try { Marshal.ReleaseComObject(rightPic); } catch { } rightPic = null; }
                return false;
            }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static void CollectPictureCandidates(PptShapes shapes, List<PptShape> list, bool searchGroups)
        {
            if (shapes == null || list == null) return;
            int count = shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = shapes[i];
                    CollectPictureCandidatesFromShape(sh, list, searchGroups);
                    sh = null;
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
        }

        private static void CollectPictureCandidatesFromGroup(Microsoft.Office.Interop.PowerPoint.GroupShapes group, List<PptShape> list, bool searchGroups)
        {
            if (group == null || list == null) return;
            int count = group.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = group[i];
                    CollectPictureCandidatesFromShape(sh, list, searchGroups);
                    sh = null;
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
        }

        private static void CollectPictureCandidatesFromShape(PptShape sh, List<PptShape> list, bool searchGroups)
        {
            if (sh == null || list == null) return;

            if (searchGroups && sh.Type == MsoShapeType.msoGroup)
            {
                Microsoft.Office.Interop.PowerPoint.GroupShapes group = null;
                try
                {
                    group = sh.GroupItems;
                    if (group != null)
                        CollectPictureCandidatesFromGroup(group, list, searchGroups);
                }
                finally
                {
                    if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
                }
                return;
            }

            if (!IsPictureCandidate(sh))
                return;

            list.Add(sh);
        }

        /// <summary>P4-6: スライド5の右側画像を右端でトリミング。</summary>
        public bool CheckTask_1_4_06()
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
                        PptShape rightmostPicture = null;
                        float maxLeft = float.MinValue;
                        int count = shapes.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (!PowerPointCheckerCommon.IsPictureShape(sh)) continue;
                                float left = (float)sh.Left;
                                if (left > maxLeft)
                                {
                                    maxLeft = left;
                                    if (rightmostPicture != null) Marshal.ReleaseComObject(rightmostPicture);
                                    rightmostPicture = sh;
                                    sh = null;
                                }
                            }
                            finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                        }
                        if (rightmostPicture == null) return false;
                        try
                        {
                            dynamic pf = rightmostPicture.PictureFormat;
                            float cr = (float)pf.CropRight;
                            Marshal.ReleaseComObject(pf);
                            return cr > 0.1;
                        }
                        finally { if (rightmostPicture != null) Marshal.ReleaseComObject(rightmostPicture); }
                    }
                    finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
                }
                finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }
            }
            catch { return false; }
            finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
        }

        private const float Task4_7FillBrightnessTarget = 0.6f;
        private const float Task4_7FillBrightnessTolerance = 0.09f;
        private const float Task4_7FillRgbLuminanceMin = 0.70f;
        private const float Task4_7FillRgbLuminanceMax = 0.82f;
        private const float Task4_7LineWeightTarget = 0.75f;
        private const float Task4_7LineWeightTolerance = 0.2f;

        /// <summary>P4-7: スライド2のテキストボックスに塗りつぶし（アクセント1・白+基本色60％）・枠線（濃い青）0.75pt。</summary>
        public bool CheckTask_1_4_07()
        {
            Presentation pres = null;
            string tempPath = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 2);
                    if (slide == null) return false;

                    if (TryValidateTask4_7OnSlideCom(slide))
                        return true;

                    tempPath = Path.Combine(Path.GetTempPath(), "mos_4_7_check_" + Guid.NewGuid().ToString("N") + ".pptx");
                    pres.SaveCopyAs(tempPath);
                    return PptxTask4_7StyleOpenXmlReader.TryValidateOnSlide(tempPath, 2);
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
                if (tempPath != null && File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
        }

        private static bool TryValidateTask4_7OnSlideCom(Slide slide)
        {
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
                        if (TryValidateTask4_7OnTextShapeCom(sh))
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

        private static bool TryValidateTask4_7OnTextShapeCom(PptShape sh)
        {
            if (sh == null || sh.HasTextFrame != MsoTriState.msoTrue)
                return false;

            string text = "";
            try
            {
                var tf = sh.TextFrame;
                if (tf != null) text = tf.TextRange?.Text ?? "";
            }
            catch { }
            if (string.IsNullOrWhiteSpace(text))
                return false;

            Microsoft.Office.Interop.PowerPoint.FillFormat fill = null;
            Microsoft.Office.Interop.PowerPoint.LineFormat line = null;
            try
            {
                fill = sh.Fill;
                if (fill == null || fill.Visible != MsoTriState.msoTrue)
                    return false;

                line = sh.Line;
                if (line == null || line.Visible != MsoTriState.msoTrue)
                    return false;

                float weight = 0f;
                try { weight = (float)line.Weight; } catch { }
                if (Math.Abs(weight - Task4_7LineWeightTarget) > Task4_7LineWeightTolerance)
                    return false;

                Microsoft.Office.Interop.PowerPoint.ColorFormat fillColor = null;
                Microsoft.Office.Interop.PowerPoint.ColorFormat lineColor = null;
                try
                {
                    fillColor = fill.ForeColor;
                    lineColor = line.ForeColor;
                    if (fillColor == null || lineColor == null)
                        return false;

                    return IsTask4_7FillColor(fillColor, out _)
                        && IsTask4_7LineColor(lineColor, out _);
                }
                finally
                {
                    if (lineColor != null) { try { Marshal.ReleaseComObject(lineColor); } catch { } }
                    if (fillColor != null) { try { Marshal.ReleaseComObject(fillColor); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (line != null) { try { Marshal.ReleaseComObject(line); } catch { } }
                if (fill != null) { try { Marshal.ReleaseComObject(fill); } catch { } }
            }
        }

        private static bool IsTask4_7FillColor(
            Microsoft.Office.Interop.PowerPoint.ColorFormat cf, out string failReason)
        {
            failReason = null;
            if (cf == null)
            {
                failReason = "null";
                return false;
            }

            try
            {
                if (cf.ObjectThemeColor != MsoThemeColorIndex.msoThemeColorAccent1)
                {
                    failReason = "theme is not Accent1";
                    return false;
                }
            }
            catch
            {
                failReason = "cannot read theme";
                return false;
            }

            if (TryGetThemeColorAdjustment(cf, out float adjustment))
            {
                if (Math.Abs(adjustment - Task4_7FillBrightnessTarget) <= Task4_7FillBrightnessTolerance)
                    return true;

                // COM が Brightness=0 のまま返す環境のみ: 60% 相当の淡い RGB で補完
                if (adjustment <= 0.15f && TryIsTask4_7FillRgb60Percent(cf))
                    return true;

                failReason = $"brightness={adjustment:F4} (expected ~{Task4_7FillBrightnessTarget} ±{Task4_7FillBrightnessTolerance})";
                return false;
            }

            if (TryIsTask4_7FillRgb60Percent(cf))
                return true;

            failReason = "no brightness and RGB not light accent blue";
            return false;
        }

        private static bool IsTask4_7LineColor(
            Microsoft.Office.Interop.PowerPoint.ColorFormat cf, out string failReason)
        {
            failReason = null;
            if (cf == null)
            {
                failReason = "null";
                return false;
            }

            if (TryIsTask4_7LineDarkBlueRgb(cf))
                return true;

            failReason = "not 濃い青 (#002060 付近)";
            return false;
        }

        private static bool TryGetThemeColorAdjustment(Microsoft.Office.Interop.PowerPoint.ColorFormat cf, out float adjustment)
        {
            adjustment = 0f;
            if (cf == null) return false;
            try
            {
                adjustment = (float)cf.Brightness;
                return true;
            }
            catch { }

            try
            {
                dynamic dynamicCf = cf;
                adjustment = (float)dynamicCf.TintAndShade;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>アクセント1・白+基本色60％相当の RGB（正答例 #B4C7E7 付近）。</summary>
        private static bool TryIsTask4_7FillRgb60Percent(Microsoft.Office.Interop.PowerPoint.ColorFormat cf)
        {
            if (!TryGetRgbComponents(cf, out int r, out int g, out int b))
                return false;

            double lum = GetRelativeLuminance(r, g, b);
            if (lum < Task4_7FillRgbLuminanceMin || lum > Task4_7FillRgbLuminanceMax)
                return false;
            return b >= 150 && g >= 130 && r >= 160 && r <= 200;
        }

        private static bool TryIsTask4_7LineDarkBlueRgb(Microsoft.Office.Interop.PowerPoint.ColorFormat cf)
        {
            if (!TryGetRgbComponents(cf, out int r, out int g, out int b))
                return false;
            return Task4_7LineColorRules.IsDarkBlueRgb(r, g, b);
        }

        private static bool TryGetRgbComponents(Microsoft.Office.Interop.PowerPoint.ColorFormat cf, out int r, out int g, out int b)
        {
            r = g = b = 0;
            if (cf == null) return false;
            try
            {
                int rgb = (int)cf.RGB;
                r = rgb & 0xFF;
                g = (rgb >> 8) & 0xFF;
                b = (rgb >> 16) & 0xFF;
                return true;
            }
            catch { return false; }
        }

        private static double GetRelativeLuminance(int r, int g, int b)
        {
            return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0;
        }

        private const float Task4_8VerticalCenterTolerance = 5.0f;

        /// <summary>P4-8: スライド3のコンテンツ領域テキストボックスを垂直方向中央に配置。</summary>
        public bool CheckTask_1_4_08()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                PageSetup ps = null;
                Slide slide = null;
                var candidates = new List<PptShape>();
                try
                {
                    ps = pres.PageSetup;
                    float slideHeight = ps.SlideHeight;
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 3);
                    if (slide == null) return false;
                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
                        if (shapes == null) return false;
                        CollectTask4_8TextShapeCandidates(shapes, candidates, searchGroups: true);
                        foreach (PptShape candidate in candidates)
                        {
                            if (IsShapeVerticallyCenteredOnSlide(candidate, slideHeight))
                                return true;
                        }
                        return false;
                    }
                    finally { if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } } }
                }
                finally
                {
                    foreach (PptShape candidate in candidates)
                    {
                        if (candidate != null) { try { Marshal.ReleaseComObject(candidate); } catch { } }
                    }
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                    if (ps != null) { try { Marshal.ReleaseComObject(ps); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        private static void CollectTask4_8TextShapeCandidates(PptShapes shapes, List<PptShape> list, bool searchGroups)
        {
            if (shapes == null || list == null) return;
            int count = shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = shapes[i];
                    CollectTask4_8TextShapeFromShape(sh, list, searchGroups);
                    sh = null;
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
        }

        private static void CollectTask4_8TextShapeFromGroup(Microsoft.Office.Interop.PowerPoint.GroupShapes group, List<PptShape> list, bool searchGroups)
        {
            if (group == null || list == null) return;
            int count = group.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = group[i];
                    CollectTask4_8TextShapeFromShape(sh, list, searchGroups);
                    sh = null;
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
        }

        private static void CollectTask4_8TextShapeFromShape(PptShape sh, List<PptShape> list, bool searchGroups)
        {
            if (sh == null || list == null) return;

            if (searchGroups && sh.Type == MsoShapeType.msoGroup)
            {
                Microsoft.Office.Interop.PowerPoint.GroupShapes group = null;
                try
                {
                    group = sh.GroupItems;
                    if (group != null)
                        CollectTask4_8TextShapeFromGroup(group, list, searchGroups);
                }
                finally
                {
                    if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
                }
                return;
            }

            if (sh.HasTextFrame != MsoTriState.msoTrue)
                return;

            string text = "";
            try
            {
                var tf = sh.TextFrame;
                if (tf != null) text = tf.TextRange?.Text ?? "";
            }
            catch { }

            if (!IsTask4_8TargetTextShape(sh, text))
                return;

            list.Add(sh);
        }

        private static bool IsTask4_8TargetTextShape(PptShape sh, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;
            if (IsTask4_8ExcludedTextShape(sh, text))
                return false;
            return IsTask4_8BulletBodyText(text);
        }

        private static bool IsTask4_8BulletBodyText(string text)
        {
            return text.IndexOf('•') >= 0
                || text.IndexOf('・') >= 0
                || text.IndexOf('\r') >= 0
                || text.IndexOf('\n') >= 0;
        }

        private static bool IsTask4_8ExcludedTextShape(PptShape sh, string text)
        {
            string trimmed = text.Trim();
            if (trimmed.Equals("教育理念", StringComparison.Ordinal))
                return true;
            if (IsTask4_8TitlePlaceholder(sh))
                return true;
            return false;
        }

        private static bool IsTask4_8TitlePlaceholder(PptShape sh)
        {
            if (sh == null || sh.Type != MsoShapeType.msoPlaceholder)
                return false;

            PlaceholderFormat pf = null;
            try
            {
                pf = sh.PlaceholderFormat;
                if (pf == null) return false;
                var pt = (PpPlaceholderType)pf.Type;
                return pt == PpPlaceholderType.ppPlaceholderTitle
                    || pt == PpPlaceholderType.ppPlaceholderCenterTitle
                    || pt == PpPlaceholderType.ppPlaceholderVerticalTitle
                    || pt == PpPlaceholderType.ppPlaceholderSubtitle;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (pf != null) { try { Marshal.ReleaseComObject(pf); } catch { } }
            }
        }

        private static bool IsShapeVerticallyCenteredOnSlide(PptShape sh, float slideHeight)
        {
            if (sh == null)
                return false;
            float shapeCenter = (float)sh.Top + ((float)sh.Height / 2.0f);
            float slideCenter = slideHeight / 2.0f;
            return Math.Abs(shapeCenter - slideCenter) <= Task4_8VerticalCenterTolerance;
        }

        private static PptShape GetFirstPictureOnSlide(Slide slide)
        {
            if (slide == null) return null;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return null;
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        bool isPic = IsPictureCandidate(sh);
                        if (isPic)
                        {
                            PptShape result = sh;
                            sh = null;
                            return result;
                        }
                    }
                    finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                }
                return null;
            }
            finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
        }
    }
}
