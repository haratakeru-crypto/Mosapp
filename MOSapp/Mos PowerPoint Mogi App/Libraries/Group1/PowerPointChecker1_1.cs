using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using Libraries;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_1
    {
        // UiTestAppBarWindow から呼び出されるスライド削除監視のダミー定義
        private static List<int> _previousSlideIds = new List<int>();
        private static string _previousPresentationKey = null;
        public static bool Task4PassedByThirdSlideDeletion { get; private set; }
        public static void CheckSlideDeletion(List<int> currentSlideIds) { }
        public static void CheckSlideDeletion(List<int> currentSlideIds, string presentationKey = null) { }
        public static void ResetTask4SlideDeletionState()
        {
            _previousSlideIds.Clear();
            Task4PassedByThirdSlideDeletion = false;
            _previousPresentationKey = null;
        }

        // =========================================================
        // スライド追加に伴うスライド番号補正ヘルパー
        // =========================================================
        private int GetAdjustedSlideNumber(Presentation pres, int originalNumber)
        {
            if (pres == null) return originalNumber;

            // スライド1の後ろにサマリーズームが挿入された場合、スライド1より後ろの指定スライドはインデックスを+1する
            if (originalNumber > 1 && PPLogReader.HasTask1_8SummaryZoomExecutedGlobally())
            {
                return originalNumber + 1;
            }
            return originalNumber;
        }

        // =========================================================
        // 1-1: スライド4に、レイアウト「テーブルスライド」のスライドを挿入します。
        // =========================================================
        public bool CheckTask_1_1_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null)
                {
                    Debug.WriteLine("[Task1-1] COM: ActivePresentation is null");
                    return false;
                }

                const int insertAt = 4;
                bool task18Evidence = PPLogReader.HasTask1_8SummaryZoomExecutedGlobally();
                int offsetAfterSlide1 = task18Evidence ? 1 : 0;
                int insertedSlideNum = insertAt + offsetAfterSlide1;
                Debug.WriteLine($"[Task1-1] COM: task18Evidence={task18Evidence} offsetAfterSlide1={offsetAfterSlide1} insertedSlideNum={insertedSlideNum}");

                string baselineSource;
                if (!TryGetTask1_1Baseline(pres, out List<int> baselineSlideIds, out List<string> baselineSlideNames, out baselineSource))
                {
                    Debug.WriteLine("[Task1-1] COM: FAIL baseline not available (Initial/snapshot)");
                    return false;
                }

                int baselineCount = baselineSlideIds != null ? baselineSlideIds.Count : baselineSlideNames.Count;
                int expectedSlideCount = baselineCount + 1 + offsetAfterSlide1;
                Debug.WriteLine($"[Task1-1] COM: baselineSource={baselineSource} baselineCount={baselineCount} expectedSlideCount={expectedSlideCount}");
                if (baselineSlideIds != null)
                    Debug.WriteLine($"[Task1-1] COM: baselineIds=[{string.Join(",", baselineSlideIds)}]");

                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    int actualCount = slides?.Count ?? -1;
                    Debug.WriteLine($"[Task1-1] COM: actualSlideCount={actualCount}");
                    if (slides == null || slides.Count != expectedSlideCount)
                    {
                        Debug.WriteLine($"[Task1-1] COM: FAIL slide count expected={expectedSlideCount} actual={actualCount}");
                        return false;
                    }

                    string layoutName;
                    bool tableOk = IsTableLayoutSlide(slides, insertedSlideNum, out layoutName);
                    Debug.WriteLine($"[Task1-1] COM: tableSlide[{insertedSlideNum}] layout=\"{layoutName ?? ""}\" result={(tableOk ? "PASS" : "FAIL")}");
                    if (!tableOk)
                        return false;

                    if (baselineSlideIds != null)
                    {
                        var baselineIdSet = new HashSet<int>(baselineSlideIds);
                        if (!VerifyBaselineSlidesById(slides, baselineSlideIds, insertAt, offsetAfterSlide1))
                            return false;

                        Slide insertedSlide = null;
                        try
                        {
                            insertedSlide = slides[insertedSlideNum];
                            int insertedId;
                            bool idOk = TryGetSlideId(insertedSlide, out insertedId) && !baselineIdSet.Contains(insertedId);
                            Debug.WriteLine($"[Task1-1] COM: insertedSlide[{insertedSlideNum}] id={insertedId} inBaseline={baselineIdSet.Contains(insertedId)} result={(idOk ? "PASS" : "FAIL")}");
                            if (!idOk)
                                return false;
                        }
                        finally
                        {
                            if (insertedSlide != null) { try { Marshal.ReleaseComObject(insertedSlide); } catch { } }
                        }
                    }
                    else
                    {
                        if (!VerifyBaselineSlidesByName(slides, baselineSlideNames, insertAt, offsetAfterSlide1))
                            return false;
                    }

                    Debug.WriteLine("[Task1-1] CheckTask_1_1_01: result=PASS");
                    return true;
                }
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[Task1-1] COM: exception " + ex.Message);
                return false;
            }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        private static bool TryGetTask1_1Baseline(
            Presentation pres,
            out List<int> baselineSlideIds,
            out List<string> baselineSlideNames,
            out string baselineSource)
        {
            baselineSlideIds = null;
            baselineSlideNames = null;
            baselineSource = null;

            string activePath = null;
            try { activePath = pres.FullName; } catch { }
            Debug.WriteLine($"[Task1-1] COM: activePath=\"{activePath ?? ""}\"");

            string initialPath;
            if (PowerPointCheckerCommon.TryResolveInitialPptxPath(activePath, out initialPath)
                && PptxSlideBaselineReader.TryReadOrderedSlideIds(initialPath, out baselineSlideIds))
            {
                baselineSource = "Initial:" + initialPath;
                return true;
            }

            Debug.WriteLine($"[Task1-1] COM: Initial baseline unavailable (resolved=\"{initialPath ?? ""}\")");

            PPLogReader.PPTaskSnapshotData snapshot;
            if (PPLogReader.TryLoadTaskSnapshot(1, 1, out snapshot)
                && snapshot.SlideNames != null
                && snapshot.SlideNames.Count == snapshot.SlidesCount
                && snapshot.SlidesCount >= 3)
            {
                baselineSlideNames = snapshot.SlideNames;
                baselineSource = $"Snapshot:SlidesCount={snapshot.SlidesCount}";
                return true;
            }

            Debug.WriteLine("[Task1-1] COM: snapshot baseline unavailable (missing, ID mismatch, or invalid)");
            return false;
        }

        private static int MapBaselineSlidePositionToCurrent(int snapPos1Based, int insertAt1Based, int offsetAfterSlide1)
        {
            if (snapPos1Based < insertAt1Based)
                return snapPos1Based + (snapPos1Based >= 2 ? offsetAfterSlide1 : 0);
            return snapPos1Based + 1 + offsetAfterSlide1;
        }

        private static bool VerifyBaselineSlidesById(
            Slides slides,
            IList<int> baselineSlideIds,
            int insertAt1Based,
            int offsetAfterSlide1)
        {
            for (int i = 0; i < baselineSlideIds.Count; i++)
            {
                int snapPos = i + 1;
                int currentPos = MapBaselineSlidePositionToCurrent(snapPos, insertAt1Based, offsetAfterSlide1);
                int expectedId = baselineSlideIds[i];
                Slide slide = null;
                try
                {
                    slide = slides[currentPos];
                    int slideId;
                    if (!TryGetSlideId(slide, out slideId) || slideId != expectedId)
                    {
                        Debug.WriteLine($"[Task1-1] COM: slideMap snapPos={snapPos} currentPos={currentPos} expectedId={expectedId} actualId={slideId} result=FAIL");
                        return false;
                    }
                    Debug.WriteLine($"[Task1-1] COM: slideMap snapPos={snapPos} currentPos={currentPos} expectedId={expectedId} actualId={slideId} result=PASS");
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            return true;
        }

        private static bool VerifyBaselineSlidesByName(
            Slides slides,
            IList<string> baselineSlideNames,
            int insertAt1Based,
            int offsetAfterSlide1)
        {
            for (int i = 0; i < baselineSlideNames.Count; i++)
            {
                int snapPos = i + 1;
                int currentPos = MapBaselineSlidePositionToCurrent(snapPos, insertAt1Based, offsetAfterSlide1);
                string expectedName = baselineSlideNames[i];
                Slide slide = null;
                try
                {
                    slide = slides[currentPos];
                    string slideName;
                    if (!TryGetSlideName(slide, out slideName)
                        || !string.Equals(slideName, expectedName, StringComparison.Ordinal))
                    {
                        Debug.WriteLine($"[Task1-1] COM: nameMap snapPos={snapPos} currentPos={currentPos} expected=\"{expectedName}\" actual=\"{slideName ?? ""}\" result=FAIL");
                        return false;
                    }
                    Debug.WriteLine($"[Task1-1] COM: nameMap snapPos={snapPos} currentPos={currentPos} expected=\"{expectedName}\" actual=\"{slideName}\" result=PASS");
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            return true;
        }

        private static bool IsTableLayoutSlide(Slides slides, int slideNumber, out string layoutName)
        {
            layoutName = null;
            Slide slide = null;
            CustomLayout layout = null;
            try
            {
                slide = slides[slideNumber];
                if (slide == null) return false;
                layout = slide.CustomLayout;
                if (layout == null) return false;
                try { layoutName = layout.Name ?? ""; } catch { return false; }
                return layoutName.IndexOf("テーブルスライド", StringComparison.OrdinalIgnoreCase) >= 0
                    || layoutName.IndexOf("表スライド", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch { return false; }
            finally
            {
                if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } }
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        private static bool TryGetSlideId(Slide slide, out int slideId)
        {
            slideId = 0;
            try
            {
                slideId = slide.SlideID;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetSlideName(Slide slide, out string slideName)
        {
            slideName = null;
            try
            {
                slideName = slide.Name ?? "";
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>1-3: テキスト１スライド上で文字入力対象となりうるプレースホルダーか（タイトル・フッター等は除外）。</summary>
        private static bool IsTask1_3TextPlaceholderCandidate(PptShape sh)
        {
            if (sh == null) return false;
            if (sh.HasTextFrame != MsoTriState.msoTrue) return false;
            if (sh.Type != MsoShapeType.msoPlaceholder) return false;

            PlaceholderFormat pf = null;
            try
            {
                pf = sh.PlaceholderFormat;
                if (pf == null) return false;
                var pt = (PpPlaceholderType)pf.Type;
                if (IsTask1_3ExcludedPlaceholder(pt) || IsTask1_3TitlePlaceholder(pt))
                    return false;
                return IsTask1_3ContentPlaceholderType(pt);
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

        private static bool IsTask1_3ExcludedPlaceholder(PpPlaceholderType pt)
        {
            return pt == PpPlaceholderType.ppPlaceholderFooter
                || pt == PpPlaceholderType.ppPlaceholderDate
                || pt == PpPlaceholderType.ppPlaceholderSlideNumber
                || pt == PpPlaceholderType.ppPlaceholderHeader;
        }

        private static bool IsTask1_3TitlePlaceholder(PpPlaceholderType pt)
        {
            return pt == PpPlaceholderType.ppPlaceholderTitle
                || pt == PpPlaceholderType.ppPlaceholderCenterTitle
                || pt == PpPlaceholderType.ppPlaceholderVerticalTitle;
        }

        private static bool IsTask1_3ContentPlaceholderType(PpPlaceholderType pt)
        {
            return pt == PpPlaceholderType.ppPlaceholderBody
                || pt == PpPlaceholderType.ppPlaceholderVerticalBody
                || pt == PpPlaceholderType.ppPlaceholderSubtitle;
        }

        private static bool ShapeContainsText(PptShape sh, string searchText)
        {
            if (sh == null || string.IsNullOrEmpty(searchText)) return false;
            try
            {
                if (sh.HasTextFrame != MsoTriState.msoTrue) return false;
                var tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                string text = tf?.TextRange?.Text ?? "";
                return text.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        // =========================================================
        // 1-2: スライド4を非表示にします。
        // =========================================================
        public bool CheckTask_1_1_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    int slideNum = GetAdjustedSlideNumber(pres, 4);
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
                    if (slide == null) return false;
                    try
                    {
                        return slide.SlideShowTransition.Hidden == MsoTriState.msoTrue;
                    }
                    catch { return false; }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        // =========================================================
        // 1-3: スライド5のレイアウトを「テキスト１スライド」に変更します。
        //       上側プレースホルダーに「英語教育を始めたばかりのケース」と入力。
        // =========================================================
        public bool CheckTask_1_1_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                CustomLayout layout = null;
                try
                {
                    int slideNum = GetAdjustedSlideNumber(pres, 5);
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
                    if (slide == null) return false;

                    // レイアウト名チェック
                    bool layoutOk = false;
                    try
                    {
                        layout = slide.CustomLayout;
                        if (layout != null)
                        {
                            string name = layout.Name ?? "";
                            layoutOk = name.IndexOf("テキスト１スライド", StringComparison.OrdinalIgnoreCase) >= 0
                                    || name.IndexOf("テキスト1スライド", StringComparison.OrdinalIgnoreCase) >= 0
                                    || name.IndexOf("テキスト スライド", StringComparison.OrdinalIgnoreCase) >= 0;
                        }
                    }
                    catch { }
                    finally { if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } layout = null; } }

                    // 上側プレースホルダー（本文系・Top 最小）に指定文字列があること
                    const string requiredText = "英語教育を始めたばかりのケース";
                    bool textOk = false;
                    PptShapes shapes = null;
                    PptShape upperPlaceholder = null;
                    try
                    {
                        shapes = slide.Shapes;
                        if (shapes != null)
                        {
                            float minTop = float.MaxValue;
                            for (int i = 1; i <= shapes.Count; i++)
                            {
                                PptShape sh = null;
                                try
                                {
                                    sh = shapes[i];
                                    if (!IsTask1_3TextPlaceholderCandidate(sh)) continue;
                                    float top = (float)sh.Top;
                                    if (top < minTop)
                                    {
                                        if (upperPlaceholder != null)
                                        {
                                            try { Marshal.ReleaseComObject(upperPlaceholder); } catch { }
                                        }
                                        upperPlaceholder = sh;
                                        sh = null;
                                        minTop = top;
                                    }
                                }
                                catch { }
                                finally
                                {
                                    if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                                }
                            }
                            if (upperPlaceholder != null)
                                textOk = ShapeContainsText(upperPlaceholder, requiredText);
                        }
                    }
                    finally
                    {
                        if (upperPlaceholder != null) { try { Marshal.ReleaseComObject(upperPlaceholder); } catch { } }
                        if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
                    }

                    return layoutOk && textOk;
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        // =========================================================
        // 1-4: スライド6の箇条書きを2段組みに変更します。
        // =========================================================
        public bool CheckTask_1_1_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    int slideNum = GetAdjustedSlideNumber(pres, 6);
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
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
                                if (sh.HasTextFrame != MsoTriState.msoTrue) continue;
                                try
                                {
                                    var tf2 = sh.TextFrame2;
                                    if (tf2 == null) continue;
                                    var col = tf2.Column;
                                    if (col == null) continue;
                                    if (col.Number == 2) return true;
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

        /// <summary>1-5: 箇条書き用コンテンツプレースホルダーか（Body/Object、タイトル・フッター等は除外）。</summary>
        private static bool IsTask1_5BulletListPlaceholder(PptShape sh)
        {
            if (sh == null) return false;
            if (sh.HasTextFrame != MsoTriState.msoTrue) return false;
            if (sh.Type != MsoShapeType.msoPlaceholder) return false;

            PlaceholderFormat pf = null;
            try
            {
                pf = sh.PlaceholderFormat;
                if (pf == null) return false;
                var pt = (PpPlaceholderType)pf.Type;
                if (IsTask1_3ExcludedPlaceholder(pt) || IsTask1_3TitlePlaceholder(pt))
                    return false;
                return pt == PpPlaceholderType.ppPlaceholderBody
                    || pt == PpPlaceholderType.ppPlaceholderVerticalBody
                    || pt == PpPlaceholderType.ppPlaceholderObject;
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

        /// <summary>1-5: 箇条書き書式が付いた段落を含むか。</summary>
        private static bool ShapeHasBulletParagraph(PptShape sh)
        {
            if (sh == null) return false;
            Microsoft.Office.Interop.PowerPoint.TextRange tr = null;
            try
            {
                var tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                if (tf == null) return false;
                tr = tf.TextRange;
                if (tr == null) return false;

                int paraCount = tr.Paragraphs().Count;
                for (int i = 1; i <= paraCount; i++)
                {
                    Microsoft.Office.Interop.PowerPoint.TextRange para = null;
                    try
                    {
                        para = tr.Paragraphs(i, 1);
                        if (para?.ParagraphFormat?.Bullet == null) continue;
                        if (para.ParagraphFormat.Bullet.Type != PpBulletType.ppBulletNone)
                            return true;
                    }
                    catch { }
                    finally
                    {
                        if (para != null) { try { Marshal.ReleaseComObject(para); } catch { } }
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        // =========================================================
        // 1-5: スライド8の箇条書きのプレースホルダーの文字の間隔を広げます。幅を「4pt」にします。
        // =========================================================
        public bool CheckTask_1_1_05()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    int slideNum = GetAdjustedSlideNumber(pres, 8);
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
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
                                if (!IsTask1_5BulletListPlaceholder(sh)) continue;
                                if (!ShapeHasBulletParagraph(sh)) continue;
                                try
                                {
                                    var tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                    if (tf == null) continue;
                                    var tr = tf.TextRange;
                                    if (tr == null) continue;
                                    if (string.IsNullOrWhiteSpace(tr.Text)) continue;
                                    try
                                    {
                                        dynamic tr2 = sh.TextFrame2?.TextRange;
                                        if (tr2 != null)
                                        {
                                            try
                                            {
                                                float spacing = (float)tr2.Font.Spacing;
                                                if (Math.Abs(spacing - 4.0f) < 0.5f) return true;
                                            }
                                            catch { }
                                        }
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

        // =========================================================
        // 1-6: スライド8にセクションを追加します。セクション名は「まとめ」にします。
        // =========================================================
        public bool CheckTask_1_1_06()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                SectionProperties sectionProps = null;
                try
                {
                    int slideNum = GetAdjustedSlideNumber(pres, 8);
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
                    if (slide == null) return false;
                    try
                    {
                        sectionProps = pres.SectionProperties;
                        if (sectionProps == null) return false;
                        int sectionIndex = (int)slide.sectionIndex;
                        if (sectionIndex < 1) return false;
                        string name = sectionProps.Name(sectionIndex) ?? "";
                        return name.IndexOf("まとめ", StringComparison.OrdinalIgnoreCase) >= 0;
                    }
                    catch { return false; }
                }
                finally
                {
                    if (sectionProps != null) { try { Marshal.ReleaseComObject(sectionProps); } catch { } }
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        // =========================================================
        // 1-7: スライド1のセクション名を「はじめに」とします。
        // =========================================================
        public bool CheckTask_1_1_07()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                SectionProperties sectionProps = null;
                try
                {
                    int slideNum = GetAdjustedSlideNumber(pres, 1); // スライド1はズレないが、念のため適用
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
                    if (slide == null) return false;
                    try
                    {
                        sectionProps = pres.SectionProperties;
                        if (sectionProps == null) return false;
                        int sectionIndex = (int)slide.sectionIndex;
                        if (sectionIndex < 1) return false;
                        string name = sectionProps.Name(sectionIndex) ?? "";
                        return name.IndexOf("はじめに", StringComparison.OrdinalIgnoreCase) >= 0;
                    }
                    catch { return false; }
                }
                finally
                {
                    if (sectionProps != null) { try { Marshal.ReleaseComObject(sectionProps); } catch { } }
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        // =========================================================
        // 1-8: サマリーズームスライドを挿入します。
        //       スライド1の後ろ（スライド2）にタイトル「ご提案のポイント」のサマリーズームを挿入。
        //       リンク先は「1.教育理念」「4.募集要項」の2件のみ。スライド1・8へのリンクは不可。
        //       COM（タイトル・ズーム図形）と Open XML（リンク先タイトル・禁止スライド）の両方で検証。
        // =========================================================
        public bool CheckTask_1_1_08()
        {
            const int summaryZoomSlideNumber = 2;
            const string summaryTitle = "ご提案のポイント";
            const string linkTitle1 = "1.教育理念";
            const string linkTitle2 = "4.募集要項";
            const int forbiddenOriginalSlide8 = 8;
            int forbiddenSlide8Index = forbiddenOriginalSlide8 >= summaryZoomSlideNumber
                ? forbiddenOriginalSlide8 + 1
                : forbiddenOriginalSlide8;
            var forbiddenLinkIndices = new[] { 1, forbiddenSlide8Index };

            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                if (!CheckSummaryZoomSlideCom(pres, summaryZoomSlideNumber, summaryTitle))
                    return false;
                return ValidateSummaryZoomOpenXml(
                    pres,
                    summaryZoomSlideNumber,
                    linkTitle1,
                    linkTitle2,
                    forbiddenLinkIndices);
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        private static bool CheckSummaryZoomSlideCom(Presentation pres, int summaryZoomSlideNumber, string summaryTitle)
        {
            Slide slide = null;
            try
            {
                slide = PowerPointCheckerCommon.GetSlideByNumber(pres, summaryZoomSlideNumber);
                if (slide == null) return false;

                PptShapes shapes = null;
                try
                {
                    shapes = slide.Shapes;
                    if (shapes == null) return false;

                    bool titleOk = false;
                    bool hasZoomShape = false;
                    const int msoZoom = 21;

                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        PptShape sh = null;
                        try
                        {
                            sh = shapes[i];
                            string shapeName = "";
                            string shapeAlt = "";
                            int shapeType = (int)sh.Type;
                            try { shapeName = sh.Name ?? ""; } catch { }
                            try { shapeAlt = sh.AlternativeText ?? ""; } catch { }

                            if (sh.HasTextFrame == MsoTriState.msoTrue)
                            {
                                var tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                string text = tf?.TextRange?.Text ?? "";
                                if (text.IndexOf(summaryTitle, StringComparison.OrdinalIgnoreCase) >= 0)
                                    titleOk = true;
                            }

                            bool shapeIsZoom = shapeType == msoZoom
                                || shapeName.IndexOf("Zoom", StringComparison.OrdinalIgnoreCase) >= 0
                                || shapeName.IndexOf("ズーム", StringComparison.OrdinalIgnoreCase) >= 0
                                || shapeName.IndexOf("Summary", StringComparison.OrdinalIgnoreCase) >= 0
                                || shapeAlt.IndexOf("Zoom", StringComparison.OrdinalIgnoreCase) >= 0
                                || shapeAlt.IndexOf("ズーム", StringComparison.OrdinalIgnoreCase) >= 0;
                            if (shapeIsZoom)
                                hasZoomShape = true;
                        }
                        catch { }
                        finally { if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } } }
                    }

                    return titleOk && hasZoomShape;
                }
                finally { if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } } }
            }
            catch { return false; }
            finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
        }

        private static bool ValidateSummaryZoomOpenXml(
            Presentation pres,
            int summaryZoomSlideNumber,
            string linkTitle1,
            string linkTitle2,
            int[] forbiddenLinkIndices)
        {
            string originalPptxPath = null;
            try { originalPptxPath = pres.FullName; } catch { }

            string validationPptxPath = null;
            string tempPptxPath = null;
            try
            {
                tempPptxPath = Path.Combine(Path.GetTempPath(), "Mosapp_1_8_" + Guid.NewGuid().ToString("N") + ".pptx");
                try
                {
                    pres.SaveCopyAs(
                        tempPptxPath,
                        PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                        MsoTriState.msoFalse);
                    if (File.Exists(tempPptxPath))
                        validationPptxPath = tempPptxPath;
                }
                catch { }

                if (string.IsNullOrWhiteSpace(validationPptxPath))
                {
                    if (string.IsNullOrWhiteSpace(originalPptxPath) || !File.Exists(originalPptxPath))
                        return false;
                    if (!originalPptxPath.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                        return false;
                    validationPptxPath = originalPptxPath;
                }

                return PptxSlideZoomLinkReader.TryValidateSummaryZoomTargetTitles(
                    validationPptxPath,
                    summaryZoomSlideNumber,
                    linkTitle1,
                    linkTitle2,
                    forbiddenLinkIndices,
                    out _);
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(tempPptxPath))
                {
                    try { if (File.Exists(tempPptxPath)) File.Delete(tempPptxPath); } catch { }
                }
            }
        }
    }
}
