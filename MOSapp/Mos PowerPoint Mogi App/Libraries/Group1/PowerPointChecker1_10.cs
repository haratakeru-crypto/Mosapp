using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    /// <summary>プロジェクト10（P10-1〜P10-8）。Phase B でタスク単位に Legacy 移植。</summary>
    public class PowerPointChecker1_10
    {
        private const string P10_1ExpectedThemeName = "木版活字";
        private const int P10_2FirstSlideNumber = 2;
        private const int P10_2LastSlideNumber = 7;
        private const string P10_3TitleLayoutNameJa = "タイトルスライド";
        private const string P10_3TitleLayoutNameEn = "Title Slide";
        private const string P10_4TwoContentLayoutName = "２つのコンテンツ";
        private const string P10_6LayoutName = "タイトル付きの図と表";
        private const string P10_8ExpectedFooterText = "四季を楽しむ";

        /// <summary>P10-1: スライドマスターテーマ「木版活字」（旧10-5）。</summary>
        public bool CheckTask_1_10_01()
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
                                string name;
                                try { name = design.Name ?? ""; }
                                catch { return false; }
                                return name.IndexOf(P10_1ExpectedThemeName, StringComparison.OrdinalIgnoreCase) >= 0;
                            }
                            finally
                            {
                                if (design != null) { try { Marshal.ReleaseComObject(design); } catch { } }
                            }
                        }
                        finally
                        {
                            if (range != null) { try { Marshal.ReleaseComObject(range); } catch { } }
                        }
                    }
                    finally
                    {
                        if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
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

        /// <summary>P10-2: スライド2〜7にスライド番号表示（新規）。スライド1はP10-3で非表示のため対象外。</summary>
        public bool CheckTask_1_10_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                return SlidesInRangeHaveVisibleSlideNumber(pres, P10_2FirstSlideNumber, P10_2LastSlideNumber);
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        private static bool SlidesInRangeHaveVisibleSlideNumber(Presentation pres, int firstSlide, int lastSlide)
        {
            if (pres == null) return false;
            Slides slides = null;
            try
            {
                slides = pres.Slides;
                if (slides == null || slides.Count < lastSlide)
                    return false;

                for (int slideNum = firstSlide; slideNum <= lastSlide; slideNum++)
                {
                    Slide slide = null;
                    try
                    {
                        slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
                        if (slide == null || !SlideHasVisibleSlideNumber(slide))
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
            finally
            {
                if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
            }
        }

        private static bool SlideHasVisibleSlideNumber(Slide slide)
        {
            if (slide == null) return false;
            HeadersFooters hf = null;
            try
            {
                hf = slide.HeadersFooters;
                if (hf == null) return false;
                try
                {
                    return hf.SlideNumber.Visible == MsoTriState.msoTrue;
                }
                catch { return false; }
            }
            catch { return false; }
            finally
            {
                if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } }
            }
        }

        /// <summary>P10-3: タイトルスライドレイアウトのスライド番号非表示。スライド2〜7には番号表示必須（P10-2前提）。</summary>
        public bool CheckTask_1_10_03()
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
                        int count;
                        try { count = layouts.Count; }
                        catch { return false; }

                        for (int i = 1; i <= count; i++)
                        {
                            CustomLayout cl = null;
                            try
                            {
                                cl = layouts[i];
                                if (cl == null) continue;
                                string name;
                                try { name = cl.Name ?? ""; }
                                catch { continue; }
                                if (!IsTitleSlideLayoutName(name))
                                    continue;
                                if (!IsSlideNumberHiddenOnTitleLayout(cl))
                                    return false;
                                return SlidesInRangeHaveVisibleSlideNumber(pres, P10_2FirstSlideNumber, P10_2LastSlideNumber);
                            }
                            finally
                            {
                                if (cl != null) { try { Marshal.ReleaseComObject(cl); } catch { } }
                            }
                        }
                        return false;
                    }
                    finally
                    {
                        if (layouts != null) { try { Marshal.ReleaseComObject(layouts); } catch { } }
                    }
                }
                finally
                {
                    if (master != null) { try { Marshal.ReleaseComObject(master); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        private static bool IsTitleSlideLayoutName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            return LayoutNameContains(name, P10_3TitleLayoutNameJa)
                || LayoutNameContains(name, P10_3TitleLayoutNameEn);
        }

        private static bool IsTwoContentLayoutName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            return LayoutNameContains(name, P10_4TwoContentLayoutName);
        }

        private static bool LayoutNameContains(string layoutName, string expectedSubstring)
        {
            if (string.IsNullOrEmpty(layoutName) || string.IsNullOrEmpty(expectedSubstring))
                return false;
            string normalized = NormalizeLayoutNameForMatch(layoutName);
            string normalizedExpected = NormalizeLayoutNameForMatch(expectedSubstring);
            return normalized.IndexOf(normalizedExpected, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>レイアウト名の空白除去・全角数字を半角に統一して比較用に正規化。</summary>
        private static string NormalizeLayoutNameForMatch(string name)
        {
            if (string.IsNullOrEmpty(name)) return "";
            name = name
                .Replace(" ", "")
                .Replace("\u3000", "")
                .Replace("\t", "")
                .Replace("\r", "")
                .Replace("\n", "");
            for (int i = 0; i <= 9; i++)
                name = name.Replace((char)('\uFF10' + i), (char)('0' + i));
            return name;
        }

        /// <summary>11-4 と同様、プレースホルダー可視性を主判定。補助で SlideNumber.Visible == msoFalse も可。</summary>
        private static bool IsSlideNumberHiddenOnTitleLayout(CustomLayout layout)
        {
            if (layout == null) return false;

            if (!LayoutHasVisibleSlideNumberPlaceholder(layout))
                return true;

            HeadersFooters hf = null;
            try
            {
                try { hf = layout.HeadersFooters; }
                catch { return false; }
                if (hf == null) return false;
                try
                {
                    return hf.SlideNumber.Visible == MsoTriState.msoFalse;
                }
                catch { return false; }
            }
            finally
            {
                if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } }
            }
        }

        private static bool LayoutHasVisibleSlideNumberPlaceholder(CustomLayout layout)
        {
            if (layout == null) return false;
            PptShapes shapes = null;
            try
            {
                shapes = layout.Shapes;
                if (shapes == null) return false;
                int shapeCount;
                try { shapeCount = shapes.Count; }
                catch { return false; }

                for (int i = 1; i <= shapeCount; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        try { sh = shapes[i]; }
                        catch { continue; }
                        if (sh == null) continue;

                        try
                        {
                            if (sh.Type != MsoShapeType.msoPlaceholder)
                                continue;
                        }
                        catch { continue; }

                        PlaceholderFormat pf = null;
                        try
                        {
                            try { pf = sh.PlaceholderFormat; }
                            catch { continue; }
                            if (pf == null) continue;
                            if ((PpPlaceholderType)pf.Type != PpPlaceholderType.ppPlaceholderSlideNumber)
                                continue;
                            try
                            {
                                if (sh.Visible == MsoTriState.msoTrue)
                                    return true;
                            }
                            catch { return true; }
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

        /// <summary>P10-4: 「２つのコンテンツ」レイアウト背景非表示（旧10-6）。</summary>
        public bool CheckTask_1_10_04()
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
                        int count;
                        try { count = layouts.Count; }
                        catch { return false; }

                        for (int i = 1; i <= count; i++)
                        {
                            CustomLayout cl = null;
                            try
                            {
                                cl = layouts[i];
                                if (cl == null) continue;
                                string name;
                                try { name = cl.Name ?? ""; }
                                catch { continue; }
                                if (!IsTwoContentLayoutName(name))
                                    continue;
                                try
                                {
                                    return cl.DisplayMasterShapes == MsoTriState.msoFalse;
                                }
                                catch { return false; }
                            }
                            finally
                            {
                                if (cl != null) { try { Marshal.ReleaseComObject(cl); } catch { } }
                            }
                        }
                        return false;
                    }
                    finally
                    {
                        if (layouts != null) { try { Marshal.ReleaseComObject(layouts); } catch { } }
                    }
                }
                finally
                {
                    if (master != null) { try { Marshal.ReleaseComObject(master); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        /// <summary>P10-5: スライドマスターでフッターPH削除（新規）。</summary>
        public bool CheckTask_1_10_05()
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
                    return !SlideMasterHasVisibleFooterPlaceholder(master);
                }
                finally
                {
                    if (master != null) { try { Marshal.ReleaseComObject(master); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        private static bool SlideMasterHasVisibleFooterPlaceholder(Master slideMaster)
        {
            if (slideMaster == null) return false;
            PptShapes shapes = null;
            try
            {
                shapes = slideMaster.Shapes;
                if (shapes == null) return false;
                int count;
                try { count = shapes.Count; }
                catch { return false; }

                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        if (sh == null) continue;
                        try
                        {
                            if (sh.Type != MsoShapeType.msoPlaceholder)
                                continue;
                        }
                        catch { continue; }

                        PlaceholderFormat pf = null;
                        try
                        {
                            pf = sh.PlaceholderFormat;
                            if (pf == null) continue;
                            if ((PpPlaceholderType)pf.Type != PpPlaceholderType.ppPlaceholderFooter)
                                continue;
                            try
                            {
                                if (sh.Visible == MsoTriState.msoTrue)
                                    return true;
                            }
                            catch { return true; }
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

        /// <summary>P10-6: レイアウト複製「タイトル付きの図と表」（旧10-7）。</summary>
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
                        int count;
                        try { count = layouts.Count; }
                        catch { return false; }

                        for (int i = 1; i <= count; i++)
                        {
                            CustomLayout cl = null;
                            try
                            {
                                cl = layouts[i];
                                if (cl == null) continue;
                                string name;
                                try { name = cl.Name ?? ""; }
                                catch { continue; }
                                if (name.IndexOf(P10_6LayoutName, StringComparison.OrdinalIgnoreCase) < 0)
                                    continue;
                                return LayoutHasPictureLeftOfTablePlaceholders(cl);
                            }
                            finally
                            {
                                if (cl != null) { try { Marshal.ReleaseComObject(cl); } catch { } }
                            }
                        }
                        return false;
                    }
                    finally
                    {
                        if (layouts != null) { try { Marshal.ReleaseComObject(layouts); } catch { } }
                    }
                }
                finally
                {
                    if (master != null) { try { Marshal.ReleaseComObject(master); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        /// <summary>図プレースホルダーが左・表プレースホルダーが右にあること（誤差 0.5pt）。</summary>
        private static bool LayoutHasPictureLeftOfTablePlaceholders(CustomLayout layout)
        {
            if (layout == null) return false;
            PptShapes shapes = null;
            try
            {
                shapes = layout.Shapes;
                if (shapes == null) return false;
                var pictureLefts = new List<float>();
                var tableLefts = new List<float>();
                int count;
                try { count = shapes.Count; }
                catch { return false; }

                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        if (sh == null) continue;
                        try
                        {
                            if (sh.Type != MsoShapeType.msoPlaceholder)
                                continue;
                        }
                        catch { continue; }

                        PlaceholderFormat pf = null;
                        try
                        {
                            pf = sh.PlaceholderFormat;
                            if (pf == null) continue;
                            PpPlaceholderType pt = (PpPlaceholderType)pf.Type;
                            float left;
                            try { left = (float)sh.Left; }
                            catch { continue; }
                            if (pt == PpPlaceholderType.ppPlaceholderPicture)
                                pictureLefts.Add(left);
                            else if (pt == PpPlaceholderType.ppPlaceholderTable)
                                tableLefts.Add(left);
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

                if (pictureLefts.Count == 0 || tableLefts.Count == 0)
                    return false;
                float minPictureLeft = pictureLefts.Min();
                float maxTableLeft = tableLefts.Max();
                return minPictureLeft + 0.5f < maxTableLeft;
            }
            catch { return false; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        /// <summary>P10-7: 配布資料マスター日付削除（旧11-4）。</summary>
        public bool CheckTask_1_10_07()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Master handoutMaster = null;
                try
                {
                    handoutMaster = pres.HandoutMaster;
                    if (handoutMaster == null) return false;
                    return !HandoutMasterHasVisibleDatePlaceholder(handoutMaster);
                }
                finally
                {
                    if (handoutMaster != null) { try { Marshal.ReleaseComObject(handoutMaster); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        /// <summary>配布資料マスター上に日付プレースホルダーが表示されていれば true。</summary>
        private static bool HandoutMasterHasVisibleDatePlaceholder(Master handoutMaster)
        {
            if (handoutMaster == null) return false;
            PptShapes shapes = null;
            try
            {
                shapes = handoutMaster.Shapes;
                if (shapes == null) return false;
                int count;
                try { count = shapes.Count; }
                catch { return false; }

                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        if (sh == null) continue;
                        try
                        {
                            if (sh.Type != MsoShapeType.msoPlaceholder)
                                continue;
                        }
                        catch { continue; }

                        PlaceholderFormat pf = null;
                        try
                        {
                            pf = sh.PlaceholderFormat;
                            if (pf == null) continue;
                            if ((PpPlaceholderType)pf.Type != PpPlaceholderType.ppPlaceholderDate)
                                continue;
                            try
                            {
                                if (sh.Visible == MsoTriState.msoTrue)
                                    return true;
                            }
                            catch { return true; }
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

        /// <summary>P10-8: 配布資料マスターフッター「四季を楽しむ」（旧11-4）。</summary>
        public bool CheckTask_1_10_08()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Master handoutMaster = null;
                try
                {
                    handoutMaster = pres.HandoutMaster;
                    if (handoutMaster == null) return false;
                    HeadersFooters hf = null;
                    try
                    {
                        hf = handoutMaster.HeadersFooters;
                        if (hf == null) return false;
                        string footer;
                        try { footer = hf.Footer.Text ?? ""; }
                        catch { return false; }
                        return footer.IndexOf(P10_8ExpectedFooterText, StringComparison.OrdinalIgnoreCase) >= 0;
                    }
                    finally
                    {
                        if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } }
                    }
                }
                finally
                {
                    if (handoutMaster != null) { try { Marshal.ReleaseComObject(handoutMaster); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }
    }
}
