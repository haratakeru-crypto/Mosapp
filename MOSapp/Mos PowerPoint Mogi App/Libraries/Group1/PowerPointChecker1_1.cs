using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_1
    {
        /// <summary>Task 4 用: 直前の Tick 時点のスライド ID の並び。</summary>
        private static List<int> _previousSlideIds = new List<int>();
        /// <summary>Task 4 用: 直前の Tick 時点のプレゼン識別子（別プレゼン比較による誤検知を防ぐ）。</summary>
        private static string _previousPresentationKey = null;
        /// <summary>Task 4 用: 3枚目（インデックス2）のスライドが削除されたと判定した場合 true。</summary>
        public static bool Task4PassedByThirdSlideDeletion { get; private set; }

        /// <summary>Task 4 用: 現在のスライド ID 一覧で削除を検出し、3枚目が削除されていればフラグを立てる。</summary>
        public static void CheckSlideDeletion(List<int> currentSlideIds)
        {
            CheckSlideDeletion(currentSlideIds, null);
        }

        /// <summary>
        /// Task 4 用: 現在のスライド ID 一覧で削除を検出し、3枚目が削除されていればフラグを立てる。
        /// </summary>
        /// <param name="presentationKey">別プレゼン比較防止用（例: FullName）。null の場合はプレゼン切替リセットのみ行わない。</param>
        public static void CheckSlideDeletion(List<int> currentSlideIds, string presentationKey = null)
        {
            if (currentSlideIds == null) return;

            // プレゼンが切り替わった場合は比較を行わず、監視状態を初期化する（別プレゼン比較の誤検知防止）
            string key = string.IsNullOrWhiteSpace(presentationKey) ? null : presentationKey.Trim();
            if (!string.IsNullOrEmpty(key))
            {
                if (!string.IsNullOrEmpty(_previousPresentationKey) &&
                    !string.Equals(_previousPresentationKey, key, StringComparison.OrdinalIgnoreCase))
                {
                    _previousSlideIds = new List<int>(currentSlideIds);
                    _previousPresentationKey = key;
                    Task4PassedByThirdSlideDeletion = false;
                    return;
                }
                _previousPresentationKey = key;
            }

            if (_previousSlideIds.Count == 0)
            {
                _previousSlideIds = new List<int>(currentSlideIds);
                return;
            }
            if (currentSlideIds.Count >= _previousSlideIds.Count)
            {
                _previousSlideIds = new List<int>(currentSlideIds);
                return;
            }

            var deletedIds = _previousSlideIds.Except(currentSlideIds).ToList();
            foreach (int deletedId in deletedIds)
            {
                int index = _previousSlideIds.IndexOf(deletedId);
                if (index == 2)
                {
                    Task4PassedByThirdSlideDeletion = true;
                    break;
                }
            }
            _previousSlideIds = new List<int>(currentSlideIds);
        }

        /// <summary>Task 4 用: 状態をリセットする。</summary>
        public static void ResetTask4SlideDeletionState()
        {
            _previousSlideIds.Clear();
            Task4PassedByThirdSlideDeletion = false;
            _previousPresentationKey = null;
        }

        public bool CheckTask_1_1_01()
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
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 4);
                    if (slide == null) return false;
                    try
                    {
                        layout = slide.CustomLayout;
                        if (layout == null) return false;
                        string name = null;
                        try { name = layout.Name ?? ""; } catch { return false; }
                        return name.IndexOf("表スライド", StringComparison.OrdinalIgnoreCase) >= 0;
                    }
                    finally { if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } } }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_1_02()
        {
            if (Task4PassedByThirdSlideDeletion) return true;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                try { if (pres.Slides.Count == 7) return true; } catch { } // 1-2→1-3→1-4完了後の結果状態（再起動後採点用）
                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null || slides.Count < 3) return false;
                    Slide slide2 = null, slide3 = null;
                    try
                    {
                        slide2 = slides[2];
                        slide3 = slides[3];
                        CustomLayout layout2 = null, layout3 = null;
                        try
                        {
                            layout2 = slide2.CustomLayout;
                            layout3 = slide3.CustomLayout;
                            if (layout2 == null || layout3 == null) return false;
                            string name2 = layout2.Name ?? "", name3 = layout3.Name ?? "";
                            return string.Equals(name2, name3, StringComparison.OrdinalIgnoreCase);
                        }
                        finally
                        {
                            if (layout2 != null) { try { Marshal.ReleaseComObject(layout2); } catch { } }
                            if (layout3 != null) { try { Marshal.ReleaseComObject(layout3); } catch { } }
                        }
                    }
                    finally
                    {
                        if (slide2 != null) { try { Marshal.ReleaseComObject(slide2); } catch { } }
                        if (slide3 != null) { try { Marshal.ReleaseComObject(slide3); } catch { } }
                    }
                }
                finally { if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_1_03()
        {
            if (Task4PassedByThirdSlideDeletion) return true;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                try { if (pres.Slides.Count == 7) return true; } catch { } // 1-2→1-3→1-4完了後の結果状態（再起動後採点用）
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 3);
                    if (slide == null) return false;
                    try { return slide.SlideShowTransition.Hidden == MsoTriState.msoTrue; }
                    catch { return false; }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_1_04()
        {
            if (Task4PassedByThirdSlideDeletion) return true;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres != null) { try { if (pres.Slides.Count == 7) return true; } catch { } } // 1-2→1-3→1-4完了後の結果状態（再起動後採点用）
            }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
            return false;
        }

        public bool CheckTask_1_1_05()
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
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 2);
                    if (slide == null) return false;
                    try
                    {
                        layout = slide.CustomLayout;
                        if (layout == null) return false;
                        string name = null;
                        try { name = layout.Name ?? ""; } catch { return false; }
                        return name.IndexOf("表スライド", StringComparison.OrdinalIgnoreCase) >= 0;
                    }
                    finally { if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } } }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_1_06()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 3);
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

        /// <summary>1-7: スライド1の吹き出しに「教育者必見」が入っているか。</summary>
        public bool CheckTask_1_1_07()
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
    }
}
