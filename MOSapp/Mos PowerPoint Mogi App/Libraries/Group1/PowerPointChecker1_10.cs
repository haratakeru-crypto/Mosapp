using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using Libraries;

namespace Libraries.Group1
{
    public class PowerPointChecker1_10
    {
        /// <summary>10-1: ドキュメント検査の結果を検証。コメントが0件かつ、ドキュメントのプロパティと個人情報が空であること。</summary>
        public bool CheckTask_1_10_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                // コメント: 全スライドで0件であること
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
                                if (comments != null && comments.Count > 0) return false;
                            }
                            finally { if (comments != null) { try { Marshal.ReleaseComObject(comments); } catch { } } }
                        }
                        finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
                    }
                }
                finally { if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } } }

                // ドキュメントのプロパティと個人情報: 指定プロパティがすべて空であること
                try
                {
                    dynamic props = pres.BuiltInDocumentProperties;
                    if (props == null) return true;
                    string[] personalPropNames = { "Author", "Manager", "Company", "Last Author", "Title", "Subject", "Keywords", "Comments" };
                    foreach (string name in personalPropNames)
                    {
                        try
                        {
                            object val = props[name].Value;
                            string s = (val == null) ? "" : (val.ToString() ?? "").Trim();
                            if (!string.IsNullOrEmpty(s)) return false;
                        }
                        catch { /* プロパティが存在しない場合は無視 */ }
                    }
                    return true;
                }
                catch { return false; }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>10-2: 目的別スライドショー「教育」があり、現在のスライド4・5・6枚目がこの順で含まれること。</summary>
        public bool CheckTask_1_10_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                int[] expectedSlideIds = new int[3];
                for (int k = 0; k < 3; k++)
                {
                    Slide slide = null;
                    try
                    {
                        slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 4 + k);
                        if (slide == null) return false;
                        expectedSlideIds[k] = slide.SlideID;
                    }
                    finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
                }

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
                                try { name = ns.Name ?? ""; } catch { continue; }
                                if (!string.Equals((name ?? "").Trim(), "教育", StringComparison.Ordinal))
                                    continue;
                                int[] showIds = TryGetNamedSlideShowSlideIds(ns);
                                if (showIds == null || showIds.Length != 3) continue;
                                if (showIds[0] == expectedSlideIds[0]
                                    && showIds[1] == expectedSlideIds[1]
                                    && showIds[2] == expectedSlideIds[2])
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

        /// <summary>
        /// NamedSlideShow.SlideIDs の COM 配列を int 列にする。
        /// null / Missing に加え、1 始まり SafeArray の先頭パディングとしての 0 をスキップする（実スライド ID は正の整数）。
        /// </summary>
        private static int[] NormalizeSlideIdsFromComArray(Array arr)
        {
            var list = new List<int>();
            for (int i = 0; i < arr.Length; i++)
            {
                object o = arr.GetValue(i);
                if (o == null || Equals(o, Missing.Value))
                    continue;
                try
                {
                    int id = Convert.ToInt32(o);
                    if (id == 0)
                        continue;
                    list.Add(id);
                }
                catch { }
            }
            return list.ToArray();
        }

        /// <summary>NamedSlideShow.SlideIDs を int 配列に変換する（取得失敗時は null）。</summary>
        private static int[] TryGetNamedSlideShowSlideIds(NamedSlideShow namedShow)
        {
            if (namedShow == null) return null;
            object raw = null;
            try
            {
                try { raw = namedShow.SlideIDs; }
                catch { return null; }
                if (raw == null) return null;
                if (raw is Array arr)
                {
                    if (arr.Length == 0) return new int[0];
                    // PowerPoint の SlideIDs は COM 上 1 始まりのため、[0]=0 や先頭 null の 4 要素配列になることがある。
                    return NormalizeSlideIdsFromComArray(arr);
                }
                return new[] { Convert.ToInt32(raw) };
            }
            catch
            {
                return null;
            }
            finally
            {
                if (raw != null && Marshal.IsComObject(raw))
                {
                    try { Marshal.ReleaseComObject(raw); } catch { }
                }
            }
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

        /// <summary>
        /// 10-4: 表示グレースケールの証跡ログに加え、スライド1の画像が「反転させたグレースケール」(Shape.BlackWhiteMode) であること。
        /// </summary>
        public bool CheckTask_1_10_04()
        {
            if (!PPLogReader.HasTask10_4GrayscaleExecuted())
                return false;
            return Slide1HasInverseGrayscalePicture();
        }

        private static bool Slide1HasInverseGrayscalePicture()
        {
            if (TrySlide1InverseGrayscaleViaCom())
                return true;
            return TrySlide1InverseGrayscaleViaOpenXml();
        }

        /// <summary>スライド1の子供の画像候補（通常画像または画像プレースホルダー）か。</summary>
        private static bool IsSlide1ChildImageShape(PptShape shape)
        {
            if (shape == null)
                return false;

            if (PowerPointCheckerCommon.IsPictureShape(shape))
                return true;

            try
            {
                if ((int)shape.Type != (int)MsoShapeType.msoPlaceholder)
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
                return name.IndexOf("Picture", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>スライド1の画像（プレースホルダー含む）で BlackWhiteMode が反転グレースケールなら true。</summary>
        private static bool TrySlide1InverseGrayscaleViaCom()
        {
            Presentation pres = null;
            Slide slide = null;
            const int expectedInverseMode = (int)MsoBlackWhiteMode.msoBlackWhiteInverseGrayScale;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null)
                    return false;

                slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
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
                            if (!IsSlide1ChildImageShape(sh))
                                continue;

                            if ((int)sh.BlackWhiteMode == expectedInverseMode)
                                return true;
                        }
                        catch
                        {
                            // 個別図形の取得失敗はスキップ
                        }
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

        private static readonly Regex Slide1InverseBwModeRegex = new Regex(
            @"\bbwMode\s*=\s*""(?:invGray|inverseGray)""",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>スライド1の OpenXML で bwMode が invGray / inverseGray なら true。</summary>
        private static bool TrySlide1InverseGrayscaleViaOpenXml()
        {
            Presentation pres = null;
            string tempPath = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null)
                    return false;

                tempPath = Path.Combine(Path.GetTempPath(), "mos_10_4_check_" + Guid.NewGuid().ToString("N") + ".pptx");
                pres.SaveCopyAs(tempPath, PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
                if (!File.Exists(tempPath))
                    return false;

                using (var package = Package.Open(tempPath, FileMode.Open, FileAccess.Read))
                {
                    var slide1Part = package.GetParts()
                        .FirstOrDefault(p => p.Uri.OriginalString.EndsWith("/slides/slide1.xml", StringComparison.OrdinalIgnoreCase));
                    if (slide1Part == null)
                        return false;

                    string slideXml;
                    using (var reader = new StreamReader(slide1Part.GetStream()))
                    {
                        slideXml = reader.ReadToEnd();
                    }

                    if (string.IsNullOrEmpty(slideXml))
                        return false;

                    return Slide1InverseBwModeRegex.IsMatch(slideXml);
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
        /// <summary>10-7: レイアウト名に「画像付きスライド」があり、グラフ用プレースホルダーがテキスト用より左にあること（問題文どおり）。</summary>
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
                                    return LayoutHasChartLeftOfTextPlaceholders(cl);
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

        /// <summary>
        /// 「グラフ」プレースホルダー（チャート）が少なくとも1つあり、「テキスト」プレースホルダー（本文）が少なくとも1つあり、
        /// 最も左にあるチャートの左端より、最も右にある本文の左端の方が右側にあること。
        /// </summary>
        private static bool LayoutHasChartLeftOfTextPlaceholders(CustomLayout layout)
        {
            if (layout == null) return false;
            PptShapes shapes = null;
            try
            {
                shapes = layout.Shapes;
                if (shapes == null) return false;
                var chartLefts = new List<float>();
                var textLefts = new List<float>();
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
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
                            PpPlaceholderType pt = (PpPlaceholderType)pf.Type;
                            float left = (float)sh.Left;
                            if (IsChartPlaceholderType(pt))
                                chartLefts.Add(left);
                            else if (IsBodyTextPlaceholderType(pt))
                                textLefts.Add(left);
                        }
                        finally { if (pf != null) { try { Marshal.ReleaseComObject(pf); } catch { } } }
                    }
                    finally { if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } } }
                }
                if (chartLefts.Count == 0 || textLefts.Count == 0) return false;
                float minChartLeft = chartLefts.Min();
                float maxTextLeft = textLefts.Max();
                // グラフを左・テキストを右：左端で比較（誤差 0.5pt）
                return minChartLeft + 0.5f < maxTextLeft;
            }
            catch { return false; }
            finally { if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } } }
        }

        /// <summary>挿入メニュー「グラフ」に相当するプレースホルダー種別。</summary>
        private static bool IsChartPlaceholderType(PpPlaceholderType pt)
        {
            return pt == PpPlaceholderType.ppPlaceholderChart;
        }

        /// <summary>挿入メニュー「テキスト」（本文系）に相当するプレースホルダー種別。</summary>
        private static bool IsBodyTextPlaceholderType(PpPlaceholderType pt)
        {
            return pt == PpPlaceholderType.ppPlaceholderBody
                || pt == PpPlaceholderType.ppPlaceholderVerticalBody;
        }
    }
}
