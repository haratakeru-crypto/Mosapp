using System;
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

namespace Libraries.Group1
{
    /// <summary>プロジェクト8（P8-1〜P8-5）。Phase B でタスク単位に Legacy 移植。</summary>
    public class PowerPointChecker1_8
    {
        private const string ExpectedSlideTitleP8_1 = "見せる！スライドの基本ルール";

        /// <summary>P8-1: 表スタイル「中間スタイル4-アクセント4」・行交互なし（旧9-2、対象スライド・スタイル名差分）。</summary>
        public bool CheckTask_1_8_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitlePlaceholderExact(pres, ExpectedSlideTitleP8_1);
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
                                        if (!IsP8_1TargetTableStyle(styleName)) continue;
                                        if (!table.HorizBanding) return true;
                                    }
                                    finally
                                    {
                                        if (ts != null) { try { Marshal.ReleaseComObject(ts); } catch { } }
                                    }
                                }
                                finally
                                {
                                    if (table != null) { try { Marshal.ReleaseComObject(table); } catch { } }
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

        private static bool IsP8_1TargetTableStyle(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName)) return false;
            string s = styleName
                .Replace("　", "")
                .Replace(" ", "")
                .Replace("-", "")
                .Replace("－", "")
                .Replace("４", "4")
                .ToLowerInvariant();
            bool isMedium4 = s.Contains("中間スタイル4") || s.Contains("mediumstyle4");
            bool isAccent4 = s.Contains("アクセント4") || s.Contains("accent4");
            return isMedium4 && isAccent4;
        }

        private const float P8_2TintTarget = 0.8f;
        private const float P8_2TintTolerance = 0.12f;

        /// <summary>P8-2: スライド1背景「濃い緑、テキスト2、白+基本色80%」（旧6-2、スライド番号・色指定差分）。</summary>
        public bool CheckTask_1_8_02()
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
                    Microsoft.Office.Interop.PowerPoint.ShapeRange bg = null;
                    try
                    {
                        bg = slide.Background;
                        if (bg == null) return false;
                        Microsoft.Office.Interop.PowerPoint.FillFormat fill = null;
                        try
                        {
                            fill = bg.Fill;
                            if (fill == null) return false;
                            if (fill.Visible != MsoTriState.msoTrue) return false;
                            Microsoft.Office.Interop.PowerPoint.ColorFormat cf = null;
                            try
                            {
                                cf = fill.ForeColor;
                                if (cf == null) return false;
                                if (cf.ObjectThemeColor != MsoThemeColorIndex.msoThemeColorText2)
                                    return false;
                                if (!TryGetThemeColorAdjustment(cf, out float adjustment))
                                    return false;
                                return Math.Abs(adjustment - P8_2TintTarget) <= P8_2TintTolerance;
                            }
                            finally
                            {
                                if (cf != null) { try { Marshal.ReleaseComObject(cf); } catch { } }
                            }
                        }
                        finally
                        {
                            if (fill != null) { try { Marshal.ReleaseComObject(fill); } catch { } }
                        }
                    }
                    finally
                    {
                        if (bg != null) { try { Marshal.ReleaseComObject(bg); } catch { } }
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

        private const float P8_3AspectRatioTarget = 16f / 9f;
        private const float P8_3AspectRatioTolerance = 0.02f;

        /// <summary>P8-3: スライドサイズ16:9（旧11-1）。VSTO 証跡優先（P8-4で上書きされるため）＋COM フォールバック。</summary>
        public bool CheckTask_1_8_03()
        {
            if (PPLogReader.HasTask8_3SlideSize16x9Executed())
                return true;
            return TryPageSetupIs16x9ViaCom();
        }

        private static bool TryPageSetupIs16x9ViaCom()
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
                    return IsPageSetupRatio16x9(pageSetup);
                }
                catch { return false; }
                finally
                {
                    if (pageSetup != null) { try { Marshal.ReleaseComObject(pageSetup); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        private static bool IsPageSetupRatio16x9(PageSetup pageSetup)
        {
            if (pageSetup == null) return false;
            float w = (float)pageSetup.SlideWidth;
            float h = (float)pageSetup.SlideHeight;
            if (h <= 0f) return false;
            float ratio = w / h;
            return Math.Abs(ratio - P8_3AspectRatioTarget) <= P8_3AspectRatioTolerance;
        }

        private const float P8_4CmToPt = 72f / 2.54f;
        private const float P8_4ExpectedWidthCm = 27.54f;
        private const float P8_4ExpectedHeightCm = 17.46f;
        private const float P8_4DimensionTolerancePt = 1.0f;

        /// <summary>P8-4: スライド寸法 幅27.54cm×高さ17.46cm（旧9-7、寸法差分）。</summary>
        public bool CheckTask_1_8_04()
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
                    float w = (float)pageSetup.SlideWidth;
                    float h = (float)pageSetup.SlideHeight;
                    float expectedWidthPt = P8_4ExpectedWidthCm * P8_4CmToPt;
                    float expectedHeightPt = P8_4ExpectedHeightCm * P8_4CmToPt;
                    return Math.Abs(w - expectedWidthPt) < P8_4DimensionTolerancePt
                        && Math.Abs(h - expectedHeightPt) < P8_4DimensionTolerancePt;
                }
                catch { return false; }
                finally
                {
                    if (pageSetup != null) { try { Marshal.ReleaseComObject(pageSetup); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        /// <summary>P8-5: 表示グレースケール証跡＋スライド2イラスト「明るいグレースケール」（旧10-4、スライド・モード差分）。</summary>
        public bool CheckTask_1_8_05()
        {
            if (!PPLogReader.HasTask8_5GrayscaleExecuted())
                return false;
            return Slide2HasLightGrayscaleIllustration();
        }

        private static bool Slide2HasLightGrayscaleIllustration()
        {
            if (TrySlide2LightGrayscaleViaCom())
                return true;
            return TrySlide2LightGrayscaleViaOpenXml();
        }

        private static bool IsSlide2IllustrationShape(PptShape shape)
        {
            if (shape == null)
                return false;

            if (PowerPointCheckerCommon.IsPictureShape(shape))
                return true;

            try
            {
                int shapeType = (int)shape.Type;
                if (shapeType == 24) // msoGraphic（イラスト・アイコン）
                    return true;

                if (shapeType != (int)MsoShapeType.msoPlaceholder)
                    return false;

                PlaceholderFormat placeholder = null;
                try
                {
                    placeholder = shape.PlaceholderFormat;
                    if (placeholder != null
                        && placeholder.Type == PpPlaceholderType.ppPlaceholderPicture)
                    {
                        return true;
                    }
                }
                finally
                {
                    if (placeholder != null)
                    {
                        try { Marshal.ReleaseComObject(placeholder); } catch { }
                    }
                }

                string name = null;
                try { name = shape.Name ?? ""; } catch { name = ""; }
                return name.IndexOf("Picture", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("Graphic", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("イラスト", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool TrySlide2LightGrayscaleViaCom()
        {
            Presentation pres = null;
            Slide slide = null;
            const int expectedLightMode = (int)MsoBlackWhiteMode.msoBlackWhiteLightGrayScale;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null)
                    return false;

                slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 2);
                if (slide == null)
                    return false;

                PptShapes shapes = null;
                try
                {
                    shapes = slide.Shapes;
                    if (shapes == null)
                        return false;

                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        PptShape sh = null;
                        try
                        {
                            sh = shapes[i];
                            if (!IsSlide2IllustrationShape(sh))
                                continue;

                            if ((int)sh.BlackWhiteMode == expectedLightMode)
                                return true;
                        }
                        catch { }
                        finally
                        {
                            if (sh != null)
                            {
                                try { Marshal.ReleaseComObject(sh); } catch { }
                            }
                        }
                    }

                    return false;
                }
                finally
                {
                    if (shapes != null)
                    {
                        try { Marshal.ReleaseComObject(shapes); } catch { }
                    }
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                if (slide != null)
                {
                    try { Marshal.ReleaseComObject(slide); } catch { }
                }
                if (pres != null)
                {
                    try { Marshal.ReleaseComObject(pres); } catch { }
                }
            }
        }

        private static readonly Regex Slide2LightBwModeRegex = new Regex(
            @"\bbwMode\s*=\s*""(?:ltGray|lightGray)""",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static bool TrySlide2LightGrayscaleViaOpenXml()
        {
            Presentation pres = null;
            string tempPath = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null)
                    return false;

                tempPath = Path.Combine(Path.GetTempPath(), "mos_8_5_check_" + Guid.NewGuid().ToString("N") + ".pptx");
                pres.SaveCopyAs(tempPath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
                if (!File.Exists(tempPath))
                    return false;

                using (var package = Package.Open(tempPath, FileMode.Open, FileAccess.Read))
                {
                    var slide2Part = package.GetParts()
                        .FirstOrDefault(p => p.Uri.OriginalString.EndsWith("/slides/slide2.xml", StringComparison.OrdinalIgnoreCase));
                    if (slide2Part == null)
                        return false;

                    string slideXml;
                    using (var reader = new StreamReader(slide2Part.GetStream()))
                    {
                        slideXml = reader.ReadToEnd();
                    }

                    if (string.IsNullOrEmpty(slideXml))
                        return false;

                    return Slide2LightBwModeRegex.IsMatch(slideXml);
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                if (pres != null)
                {
                    try { Marshal.ReleaseComObject(pres); } catch { }
                }
                if (tempPath != null && File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
        }
    }
}
