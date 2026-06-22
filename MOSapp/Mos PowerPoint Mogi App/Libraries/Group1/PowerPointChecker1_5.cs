using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using PptGroupShapes = Microsoft.Office.Interop.PowerPoint.GroupShapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_5
    {
        private const int MsoShape5pointStar = 92;
        private const int MsoShape8pointStar = 58;
        private const int MsoShape16pointStar = 59;
        private const int MsoShape24pointStar = 60;
        private const int MsoShape32pointStar = 61;

        private const float ShapeSizeTolerance = 0.5f;
        private const float PositionTolerance = 2.0f;

        /// <summary>P5-1: スライド3の丸4個の右端揃え（旧4-5）。円候補ちょうど4個・右端一致。</summary>
        public bool CheckTask_1_5_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slide slide3 = null;
                try
                {
                    slide3 = PowerPointCheckerCommon.GetSlideByNumber(pres, 3);
                    if (slide3 == null) return false;

                    List<float> rightEdges = CollectOvalRightEdgesOnSlide(slide3);
                    if (rightEdges.Count != 4) return false;

                    rightEdges.Sort();
                    float minRight = rightEdges[0];
                    float maxRight = rightEdges[rightEdges.Count - 1];
                    return Math.Abs(maxRight - minRight) <= PositionTolerance;
                }
                finally
                {
                    if (slide3 != null) { try { Marshal.ReleaseComObject(slide3); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P5-2: スライド5の小さい四角の幅を他の四角と揃える（旧5-4）。四角候補3つ以上・幅一致。</summary>
        public bool CheckTask_1_5_02()
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

                    List<float> rectangleWidths = CollectRectangleWidthsOnSlide(slide5);
                    if (rectangleWidths.Count < 3) return false;

                    rectangleWidths.Sort();
                    float minWidth = rectangleWidths[0];
                    float maxWidth = rectangleWidths[rectangleWidths.Count - 1];
                    return Math.Abs(maxWidth - minWidth) < ShapeSizeTolerance;
                }
                finally
                {
                    if (slide5 != null) { try { Marshal.ReleaseComObject(slide5); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
        /// <summary>
        /// P5-3: スライド4の星の図形をスマイルに変更（旧5-3）。
        /// スライド4にスマイル1つ・星0、プレゼン全体のスマイルも1つのみ。
        /// </summary>
        public bool CheckTask_1_5_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slide slide4 = null;
                try
                {
                    slide4 = PowerPointCheckerCommon.GetSlideByNumber(pres, 4);
                    if (slide4 == null) return false;

                    int smileysOnSlide4 = CountMatchingAutoShapesOnSlide(slide4, IsSmileyAutoShape);
                    if (smileysOnSlide4 != 1) return false;

                    int starsOnSlide4 = CountMatchingAutoShapesOnSlide(slide4, IsStarAutoShape);
                    if (starsOnSlide4 != 0) return false;

                    return CountSmileysInPresentation(pres) == 1;
                }
                finally
                {
                    if (slide4 != null) { try { Marshal.ReleaseComObject(slide4); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P5-4: スライド6の重なり順（旧4-6）。手前から「対策講座」「PC教室」「通信講座」。</summary>
        public bool CheckTask_1_5_04()
        {
            const string textTaisho = "対策講座";
            const string textPc = "PC教室";
            const string textTsushin = "通信講座";

            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slide slide6 = null;
                try
                {
                    slide6 = PowerPointCheckerCommon.GetSlideByNumber(pres, 6);
                    if (slide6 == null) return false;

                    PptShape shTaisho = null;
                    PptShape shPc = null;
                    PptShape shTsushin = null;
                    try
                    {
                        shTaisho = PowerPointCheckerCommon.FindShapeWithTextDeep(slide6, textTaisho);
                        shPc = PowerPointCheckerCommon.FindShapeWithTextDeep(slide6, textPc);
                        shTsushin = PowerPointCheckerCommon.FindShapeWithTextDeep(slide6, textTsushin);
                        if (shTaisho == null || shPc == null || shTsushin == null) return false;

                        int zTaisho = shTaisho.ZOrderPosition;
                        int zPc = shPc.ZOrderPosition;
                        int zTsushin = shTsushin.ZOrderPosition;
                        return zTaisho > zPc && zPc > zTsushin;
                    }
                    finally
                    {
                        if (shTsushin != null) { try { Marshal.ReleaseComObject(shTsushin); } catch { } }
                        if (shPc != null) { try { Marshal.ReleaseComObject(shPc); } catch { } }
                        if (shTaisho != null) { try { Marshal.ReleaseComObject(shTaisho); } catch { } }
                    }
                }
                finally
                {
                    if (slide6 != null) { try { Marshal.ReleaseComObject(slide6); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P5-5: スライド6の論理積ゲート3つをグループ化（旧5-5）。3メンバー・同一AutoShape・寸法一致。</summary>
        public bool CheckTask_1_5_05()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slide slide6 = null;
                try
                {
                    slide6 = PowerPointCheckerCommon.GetSlideByNumber(pres, 6);
                    if (slide6 == null) return false;

                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide6.Shapes;
                        if (shapes == null) return false;

                        int count = shapes.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (IsThreeMemberUniformAutoShapeGroup(sh))
                                    return true;
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
                    if (slide6 != null) { try { Marshal.ReleaseComObject(slide6); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
        /// <summary>P5-6: スライド「MOSって何？」の男の子画像の代替テキスト装飾化（旧4-3）。</summary>
        public bool CheckTask_1_5_06()
        {
            const string slideTitle = "MOSって何？";

            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, slideTitle);
                    if (slide == null) return false;

                    var pictures = new List<PptShape>();
                    try
                    {
                        CollectPictureCandidatesOnSlide(slide, pictures);
                        PptShape boyPicture = GetLargestPicture(pictures);
                        if (boyPicture == null) return false;
                        return IsShapeAltTextDecorative(boyPicture);
                    }
                    finally
                    {
                        foreach (PptShape sh in pictures)
                        {
                            if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                        }
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

        /// <summary>P5-7: スライド2の[?]アイコンに濃い赤の塗りつぶし（旧11-5）。Graphic/Icon候補・RGB濃い赤。</summary>
        public bool CheckTask_1_5_07()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slide slide2 = null;
                try
                {
                    slide2 = PowerPointCheckerCommon.GetSlideByNumber(pres, 2);
                    if (slide2 == null) return false;

                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide2.Shapes;
                        if (shapes == null) return false;

                        int count = shapes.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (!IsIconOrGraphicCandidate(sh)) continue;
                                if (TryValidateDarkRedFillOnShape(sh))
                                    return true;
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
                    if (slide2 != null) { try { Marshal.ReleaseComObject(slide2); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        private static int CountSmileysInPresentation(Presentation pres)
        {
            if (pres == null) return 0;
            Slides slides = null;
            try
            {
                slides = pres.Slides;
                if (slides == null) return 0;

                int total = 0;
                int slideCount = slides.Count;
                for (int i = 1; i <= slideCount; i++)
                {
                    Slide slide = null;
                    try
                    {
                        slide = slides[i];
                        total += CountMatchingAutoShapesOnSlide(slide, IsSmileyAutoShape);
                    }
                    finally
                    {
                        if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                    }
                }
                return total;
            }
            catch { return 0; }
            finally
            {
                if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
            }
        }

        private static int CountMatchingAutoShapesOnSlide(Slide slide, Func<PptShape, bool> predicate)
        {
            if (slide == null || predicate == null) return 0;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return 0;
                return CountMatchingAutoShapesInShapes(shapes, predicate);
            }
            catch { return 0; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static int CountMatchingAutoShapesInShapes(PptShapes shapes, Func<PptShape, bool> predicate)
        {
            if (shapes == null || predicate == null) return 0;
            int total = 0;
            int count = shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = shapes[i];
                    total += CountMatchingAutoShapesOnShape(sh, predicate);
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
            return total;
        }

        private static int CountMatchingAutoShapesOnShape(PptShape sh, Func<PptShape, bool> predicate)
        {
            if (sh == null || predicate == null) return 0;

            if (sh.Type == MsoShapeType.msoGroup)
            {
                PptGroupShapes group = null;
                try
                {
                    group = sh.GroupItems;
                    return CountMatchingAutoShapesInGroup(group, predicate);
                }
                catch { return 0; }
                finally
                {
                    if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
                }
            }

            return predicate(sh) ? 1 : 0;
        }

        private static int CountMatchingAutoShapesInGroup(PptGroupShapes group, Func<PptShape, bool> predicate)
        {
            if (group == null || predicate == null) return 0;
            int total = 0;
            int count = group.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = group[i];
                    total += CountMatchingAutoShapesOnShape(sh, predicate);
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
            return total;
        }

        private static bool IsSmileyAutoShape(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                if (sh.Type != MsoShapeType.msoAutoShape) return false;
                return sh.AutoShapeType == MsoAutoShapeType.msoShapeSmileyFace;
            }
            catch { return false; }
        }

        private static bool IsStarAutoShape(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                if (sh.Type != MsoShapeType.msoAutoShape) return false;
                int t = (int)sh.AutoShapeType;
                return t == MsoShape5pointStar || t == MsoShape8pointStar || t == MsoShape16pointStar
                    || t == MsoShape24pointStar || t == MsoShape32pointStar;
            }
            catch { return false; }
        }

        private static List<float> CollectRectangleWidthsOnSlide(Slide slide)
        {
            var widths = new List<float>();
            if (slide == null) return widths;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return widths;
                CollectRectangleWidthsInShapes(shapes, widths);
                return widths;
            }
            catch { return widths; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static void CollectRectangleWidthsInShapes(PptShapes shapes, List<float> widths)
        {
            if (shapes == null || widths == null) return;
            int count = shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = shapes[i];
                    CollectRectangleWidthsOnShape(sh, widths);
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
        }

        private static void CollectRectangleWidthsOnShape(PptShape sh, List<float> widths)
        {
            if (sh == null || widths == null) return;

            if (sh.Type == MsoShapeType.msoGroup)
            {
                PptGroupShapes group = null;
                try
                {
                    group = sh.GroupItems;
                    CollectRectangleWidthsInGroup(group, widths);
                }
                catch { }
                finally
                {
                    if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
                }
                return;
            }

            if (!IsRectangleCandidate(sh)) return;
            try
            {
                widths.Add((float)sh.Width);
            }
            catch { }
        }

        private static void CollectRectangleWidthsInGroup(PptGroupShapes group, List<float> widths)
        {
            if (group == null || widths == null) return;
            int count = group.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = group[i];
                    CollectRectangleWidthsOnShape(sh, widths);
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
        }

        /// <summary>オートシェイプの角丸四角形（幅のみ比較、縦横比は問わない）。</summary>
        private static bool IsRectangleCandidate(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                if (sh.Type != MsoShapeType.msoAutoShape) return false;
                return sh.AutoShapeType == MsoAutoShapeType.msoShapeRoundedRectangle;
            }
            catch { return false; }
        }

        /// <summary>3つのオートシェイプで同一種類・寸法が揃ったグループか。</summary>
        private static bool IsThreeMemberUniformAutoShapeGroup(PptShape sh)
        {
            if (sh == null || sh.Type != MsoShapeType.msoGroup) return false;

            PptGroupShapes group = null;
            try
            {
                group = sh.GroupItems;
                if (group == null || group.Count != 3) return false;

                MsoAutoShapeType? firstType = null;
                var widths = new List<float>(3);
                var heights = new List<float>(3);

                for (int i = 1; i <= 3; i++)
                {
                    PptShape member = null;
                    try
                    {
                        member = group[i];
                        if (member == null || member.Type != MsoShapeType.msoAutoShape) return false;

                        MsoAutoShapeType memberType;
                        try
                        {
                            memberType = member.AutoShapeType;
                        }
                        catch { return false; }

                        if (!firstType.HasValue)
                            firstType = memberType;
                        else if (firstType.Value != memberType)
                            return false;

                        widths.Add((float)member.Width);
                        heights.Add((float)member.Height);
                    }
                    finally
                    {
                        if (member != null) { try { Marshal.ReleaseComObject(member); } catch { } }
                    }
                }

                return AreUniformSizes(widths) && AreUniformSizes(heights);
            }
            catch { return false; }
            finally
            {
                if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
            }
        }

        private static bool AreUniformSizes(List<float> sizes)
        {
            if (sizes == null || sizes.Count == 0) return false;
            if (sizes.Count == 1) return true;
            sizes.Sort();
            return Math.Abs(sizes[sizes.Count - 1] - sizes[0]) < ShapeSizeTolerance;
        }

        private static List<float> CollectOvalRightEdgesOnSlide(Slide slide)
        {
            var rightEdges = new List<float>();
            if (slide == null) return rightEdges;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return rightEdges;
                CollectOvalRightEdgesInShapes(shapes, rightEdges);
                return rightEdges;
            }
            catch { return rightEdges; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static void CollectOvalRightEdgesInShapes(PptShapes shapes, List<float> rightEdges)
        {
            if (shapes == null || rightEdges == null) return;
            int count = shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = shapes[i];
                    CollectOvalRightEdgesOnShape(sh, rightEdges);
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
        }

        private static void CollectOvalRightEdgesOnShape(PptShape sh, List<float> rightEdges)
        {
            if (sh == null || rightEdges == null) return;

            if (sh.Type == MsoShapeType.msoGroup)
            {
                PptGroupShapes group = null;
                try
                {
                    group = sh.GroupItems;
                    CollectOvalRightEdgesInGroup(group, rightEdges);
                }
                catch { }
                finally
                {
                    if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
                }
                return;
            }

            if (!IsOvalCandidate(sh)) return;
            try
            {
                rightEdges.Add((float)sh.Left + (float)sh.Width);
            }
            catch { }
        }

        private static void CollectOvalRightEdgesInGroup(PptGroupShapes group, List<float> rightEdges)
        {
            if (group == null || rightEdges == null) return;
            int count = group.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = group[i];
                    CollectOvalRightEdgesOnShape(sh, rightEdges);
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
        }

        private static bool IsOvalCandidate(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                if (sh.Type != MsoShapeType.msoAutoShape) return false;
                return sh.AutoShapeType == MsoAutoShapeType.msoShapeOval;
            }
            catch { return false; }
        }

        private static void CollectPictureCandidatesOnSlide(Slide slide, List<PptShape> pictures)
        {
            if (slide == null || pictures == null) return;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return;
                CollectPictureCandidatesInShapes(shapes, pictures, searchGroups: true);
            }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static void CollectPictureCandidatesInShapes(PptShapes shapes, List<PptShape> pictures, bool searchGroups)
        {
            if (shapes == null || pictures == null) return;
            int count = shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = shapes[i];
                    CollectPictureCandidatesOnShape(sh, pictures, searchGroups);
                    sh = null;
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
        }

        private static void CollectPictureCandidatesInGroup(PptGroupShapes group, List<PptShape> pictures, bool searchGroups)
        {
            if (group == null || pictures == null) return;
            int count = group.Count;
            for (int i = 1; i <= count; i++)
            {
                PptShape sh = null;
                try
                {
                    sh = group[i];
                    CollectPictureCandidatesOnShape(sh, pictures, searchGroups);
                    sh = null;
                }
                finally
                {
                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
            }
        }

        private static void CollectPictureCandidatesOnShape(PptShape sh, List<PptShape> pictures, bool searchGroups)
        {
            if (sh == null || pictures == null) return;

            if (searchGroups && sh.Type == MsoShapeType.msoGroup)
            {
                PptGroupShapes group = null;
                try
                {
                    group = sh.GroupItems;
                    if (group != null)
                        CollectPictureCandidatesInGroup(group, pictures, searchGroups);
                }
                finally
                {
                    if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
                }
                return;
            }

            if (!IsPictureCandidate(sh)) return;
            pictures.Add(sh);
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
            try
            {
                return (float)sh.Width * (float)sh.Height;
            }
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

        private const int MsoShapeTypeGraphic = 24;

        private static bool IsIconOrGraphicCandidate(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                if ((int)sh.Type == MsoShapeTypeGraphic) return true;
                string name = sh.Name ?? "";
                return name.IndexOf("Graphic", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("Icon", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch { return false; }
        }

        private static bool TryValidateDarkRedFillOnShape(PptShape sh)
        {
            if (sh == null) return false;

            Microsoft.Office.Interop.PowerPoint.FillFormat fill = null;
            Microsoft.Office.Interop.PowerPoint.ColorFormat cf = null;
            try
            {
                fill = sh.Fill;
                if (fill == null || fill.Visible != MsoTriState.msoTrue) return false;

                cf = fill.ForeColor;
                return cf != null && IsDarkRedFillColor(cf);
            }
            catch { return false; }
            finally
            {
                if (cf != null) { try { Marshal.ReleaseComObject(cf); } catch { } }
                if (fill != null) { try { Marshal.ReleaseComObject(fill); } catch { } }
            }
        }

        private static bool IsDarkRedFillColor(Microsoft.Office.Interop.PowerPoint.ColorFormat cf)
        {
            if (cf == null) return false;
            if (!TryGetRgbComponents(cf, out int r, out int g, out int b))
                return false;

            // Office 標準色「濃い赤」#C00000 (192, 0, 0) 付近のみ（ただの赤 #FF0000 は除外）
            return r >= 175 && r <= 210 && g <= 45 && b <= 45;
        }

        private static bool TryGetRgbComponents(Microsoft.Office.Interop.PowerPoint.ColorFormat cf, out int r, out int g, out int b)
        {
            r = g = b = 0;
            if (cf == null) return false;
            try
            {
                if (cf.Type != MsoColorType.msoColorTypeRGB)
                    return false;
                int rgb = (int)cf.RGB;
                r = rgb & 0xFF;
                g = (rgb >> 8) & 0xFF;
                b = (rgb >> 16) & 0xFF;
                return true;
            }
            catch { return false; }
        }
    }
}