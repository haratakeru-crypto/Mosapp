using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace MOS_PowerPoint_app
{
    /// <summary>
    /// 起動中の PowerPoint に接続し、MOS 模擬試験のタスクごとの採点を行うクラス。
    /// </summary>
    public sealed class PowerPointGrader : IDisposable
    {
        private Application _app;
        private Presentation _activePresentation;
        private bool _disposed;

        /// <summary>Task 4 用: 直前の Tick 時点のスライド ID の並び（1枚目→2枚目→…）。</summary>
        private static List<int> _previousSlideIds = new List<int>();
        /// <summary>Task 4 用: 3枚目（インデックス2）のスライドが削除されたと判定した場合 true。採点時に参照。</summary>
        public static bool Task4PassedByThirdSlideDeletion { get; private set; }

        /// <summary>
        /// 現在起動している PowerPoint インスタンスに接続する。
        /// </summary>
        /// <returns>接続に成功した場合は true、PowerPoint が起動していない等で失敗した場合は false。</returns>
        public bool Connect()
        {
            if (_app != null)
                return true;

            try
            {
                _app = (Application)Marshal.GetActiveObject("PowerPoint.Application");
                if (_app == null)
                    return false;

                try
                {
                    _activePresentation = _app.ActivePresentation;
                }
                catch
                {
                    _activePresentation = null;
                }

                return _activePresentation != null;
            }
            catch (COMException)
            {
                _app = null;
                _activePresentation = null;
                return false;
            }
        }

        /// <summary>
        /// 指定したプロジェクト・タスクの採点を行う。
        /// </summary>
        /// <param name="projectId">プロジェクト ID（1～11）。</param>
        /// <param name="taskId">タスク ID。</param>
        /// <returns>合格なら true、不合格または未実装・範囲外なら false。</returns>
        public bool GradeTask(int projectId, int taskId)
        {
            if (_activePresentation == null)
                return false;

            switch (projectId)
            {
                case 1:
                    switch (taskId)
                    {
                        case 1: return GradeProject1Task1();
                        case 2: return GradeProject1Task2();
                        case 3: return GradeProject1Task3();
                        case 4: return GradeProject1Task4();
                        case 5: return GradeProject1Task5();
                        case 6: return GradeProject1Task6();
                        case 7: return GradeProject1Task7();
                        default: return false;
                    }
                case 2:
                    switch (taskId)
                    {
                        case 1: return GradeProject2Task1();
                        case 2: return false;
                        case 3: return false;
                        case 4: return GradeProject2Task4();
                        case 5: return false;
                        case 6: return false;
                        case 7: return false;
                        default: return false;
                    }
                case 3:
                    switch (taskId)
                    {
                        case 1: return GradeProject3Task1();
                        case 2: return GradeProject3Task2();
                        case 3: return false;
                        case 4: return false;
                        default: return false;
                    }
                case 4:
                    switch (taskId)
                    {
                        case 1: return false;
                        case 2: return false;
                        case 3: return false;
                        case 4: return GradeProject4Task4();
                        case 5: return GradeProject4Task5();
                        case 6: return GradeProject4Task6();
                        default: return false;
                    }
                case 5:
                    switch (taskId)
                    {
                        case 1: return false;
                        case 2: return false;
                        case 3: return false;
                        case 4: return false;
                        case 5: return false;
                        default: return false;
                    }
                case 6:
                    switch (taskId)
                    {
                        case 1: return false;
                        case 2: return false;
                        case 3: return false;
                        case 4: return false;
                        default: return false;
                    }
                case 7:
                    switch (taskId)
                    {
                        case 1: return false;
                        case 2: return false;
                        case 3: return false;
                        case 4: return false;
                        default: return false;
                    }
                case 8:
                    switch (taskId)
                    {
                        case 1: return false;
                        case 2: return false;
                        case 3: return false;
                        case 4: return false;
                        case 5: return false;
                        default: return false;
                    }
                case 9:
                    switch (taskId)
                    {
                        case 1: return GradeProject9Task1();
                        case 2: return false;
                        case 3: return false;
                        case 4: return false;
                        case 5: return false;
                        case 6: return GradeProject9Task6();
                        case 7: return false;
                        default: return false;
                    }
                case 10:
                    switch (taskId)
                    {
                        case 1: return false;
                        case 2: return false;
                        case 3: return false;
                        case 4: return false;
                        case 5: return false;
                        case 6: return false;
                        case 7: return false;
                        default: return false;
                    }
                case 11:
                    switch (taskId)
                    {
                        case 1: return false;
                        case 2: return false;
                        case 3: return false;
                        case 4: return false;
                        case 5: return false;
                        case 6: return false;
                        case 7: return false;
                        default: return false;
                    }
                default:
                    return false;
            }
        }

        /// <summary>
        /// 指定したスライド番号（1 始まり）のスライドを取得する。
        /// </summary>
        /// <param name="pres">プレゼンテーション。</param>
        /// <param name="slideNumber">スライド番号（1 始まり）。</param>
        /// <returns>該当スライド。範囲外または取得失敗時は null。呼び出し元で Marshal.ReleaseComObject すること。</returns>
        private static Slide GetSlideByNumber(Presentation pres, int slideNumber)
        {
            if (pres == null || slideNumber < 1)
                return null;
            try
            {
                Slides slides = pres.Slides;
                if (slides == null)
                    return null;
                try
                {
                    if (slideNumber > slides.Count)
                        return null;
                    return slides[slideNumber];
                }
                finally
                {
                    if (slides != null)
                    {
                        try { Marshal.ReleaseComObject(slides); } catch { }
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 指定したスライド内で、特定のテキストを含む図形を探す（1 階層のみ）。
        /// </summary>
        /// <param name="slide">スライド。</param>
        /// <param name="searchText">検索するテキスト。</param>
        /// <returns>該当図形。見つからない場合は null。呼び出し元で Marshal.ReleaseComObject すること。</returns>
        private static PptShape FindShapeWithText(Slide slide, string searchText)
        {
            if (slide == null || string.IsNullOrEmpty(searchText))
                return null;

            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null)
                    return null;

                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        if (sh.HasTextFrame == MsoTriState.msoTrue)
                        {
                            string text = null;
                            try
                            {
                                var pptTf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                text = pptTf.TextRange.Text;
                            }
                            catch { }
                            if (text != null && text.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                PptShape result = sh;
                                sh = null;
                                return result;
                            }
                        }
                    }
                    finally
                    {
                        if (sh != null)
                        {
                            try { Marshal.ReleaseComObject(sh); } catch { }
                        }
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (shapes != null)
                {
                    try { Marshal.ReleaseComObject(shapes); } catch { }
                }
            }
        }

        /// <summary>
        /// スライドのタイトル（またはスライド内のテキスト）に指定文字列を含むスライドを取得する。
        /// </summary>
        /// <param name="pres">プレゼンテーション。</param>
        /// <param name="titlePart">検索する文字列（部分一致、大文字小文字区別なし）。</param>
        /// <returns>該当スライド。見つからなければ null。呼び出し元で Marshal.ReleaseComObject すること。</returns>
        private static Slide GetSlideByTitle(Presentation pres, string titlePart)
        {
            if (pres == null || string.IsNullOrEmpty(titlePart))
                return null;
            try
            {
                Slides slides = pres.Slides;
                if (slides == null) return null;
                try
                {
                    int count = slides.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            PptShapes shapes = null;
                            try
                            {
                                shapes = slide.Shapes;
                                if (shapes == null) continue;
                                int sc = shapes.Count;
                                for (int j = 1; j <= sc; j++)
                                {
                                    PptShape sh = null;
                                    try
                                    {
                                        sh = shapes[j];
                                        if (sh.HasTextFrame != MsoTriState.msoTrue) continue;
                                        string text = null;
                                        try
                                        {
                                            var pptTf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                            text = pptTf.TextRange.Text ?? "";
                                        }
                                        catch { continue; }
                                        if (text.IndexOf(titlePart, StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            Slide result = slide;
                                            slide = null;
                                            return result;
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
                        finally
                        {
                            if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                        }
                    }
                    return null;
                }
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }
            }
            catch
            {
                return null;
            }
        }

        // ----- Project 1: 基本操作 -----

        private bool GradeProject1Task1()
        {
            Slide slide = null;
            CustomLayout layout = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 4);
                if (slide == null) return false;
                try
                {
                    layout = slide.CustomLayout;
                    if (layout == null) return false;
                    string name = null;
                    try
                    {
                        name = layout.Name ?? "";
                    }
                    catch { return false; }
                    return name.IndexOf("表スライド", StringComparison.OrdinalIgnoreCase) >= 0;
                }
                finally
                {
                    if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        private bool GradeProject1Task2()
        {
            // スライド2を複製 → 2枚目と3枚目が同じレイアウトであることを確認
            Slides slides = null;
            try
            {
                slides = _activePresentation.Slides;
                if (slides == null || slides.Count < 3) return false;
                Slide slide2 = null;
                Slide slide3 = null;
                try
                {
                    slide2 = slides[2];
                    slide3 = slides[3];
                    CustomLayout layout2 = null;
                    CustomLayout layout3 = null;
                    try
                    {
                        layout2 = slide2.CustomLayout;
                        layout3 = slide3.CustomLayout;
                        if (layout2 == null || layout3 == null) return false;
                        string name2 = layout2.Name ?? "";
                        string name3 = layout3.Name ?? "";
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
            finally
            {
                if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
            }
        }

        private bool GradeProject1Task3()
        {
            Slide slide = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 3);
                if (slide == null) return false;
                try
                {
                    return slide.SlideShowTransition.Hidden == MsoTriState.msoTrue;
                }
                catch { return false; }
            }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        /// <summary>Task 4 用: 現在のスライド ID 一覧で削除を検出し、3枚目が削除されていればフラグを立てる。監視タイマーから呼ぶ。</summary>
        public static void CheckSlideDeletion(List<int> currentSlideIds)
        {
            if (currentSlideIds == null)
                return;
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

        /// <summary>Task 4 用: 状態をリセットする（アプリバー表示時やプロジェクト切り替え時に呼ぶ）。</summary>
        public static void ResetTask4SlideDeletionState()
        {
            _previousSlideIds.Clear();
            Task4PassedByThirdSlideDeletion = false;
        }

        private bool GradeProject1Task4()
        {
            if (Task4PassedByThirdSlideDeletion)
                return true;
            return false;
        }

        private bool GradeProject1Task5()
        {
            // スライド2のレイアウトを「表スライド」に変更
            Slide slide = null;
            CustomLayout layout = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 2);
                if (slide == null) return false;
                try
                {
                    layout = slide.CustomLayout;
                    if (layout == null) return false;
                    string name = null;
                    try
                    {
                        name = layout.Name ?? "";
                    }
                    catch { return false; }
                    return name.IndexOf("表スライド", StringComparison.OrdinalIgnoreCase) >= 0;
                }
                finally
                {
                    if (layout != null) { try { Marshal.ReleaseComObject(layout); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        private bool GradeProject1Task6()
        {
            Slide slide = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 3);
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
                                if (col.Number == 2)
                                {
                                    return true;
                                }
                            }
                            catch { /* TextFrame2 not available */ }
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
            catch { return false; }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        private bool GradeProject1Task7()
        {
            // スライド1枚目の吹き出しに「教育者必見」と入力
            Slide slide = null;
            PptShape shape = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 1);
                if (slide == null) return false;
                shape = FindShapeWithText(slide, "教育者必見");
                return shape != null;
            }
            finally
            {
                if (shape != null) { try { Marshal.ReleaseComObject(shape); } catch { } }
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        // ----- Project 2: アニメーション・画面切り替え -----

        private bool GradeProject2Task1()
        {
            Slides slides = null;
            try
            {
                slides = _activePresentation.Slides;
                if (slides == null) return false;
                int count = slides.Count;
                if (count == 0) return false;
                for (int i = 1; i <= count; i++)
                {
                    Slide slide = null;
                    try
                    {
                        slide = slides[i];
                        try
                        {
                            var effect = slide.SlideShowTransition.EntryEffect;
                            // プッシュ 右から: ppEffectPushRight
                            if (effect != PpEntryEffect.ppEffectPushRight) return false;
                        }
                        catch { return false; }
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

        private static PptShape Find3DModelShape(Slide slide)
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
                        try
                        {
                            // msoLinked3DModel=31, mso3DModel=30
                            var st = (MsoShapeType)sh.Type;
                            if ((int)st == 31 || (int)st == 30) { PptShape r = sh; sh = null; return r; }
                        }
                        catch { }
                        try
                        {
                            if (sh.Name != null && sh.Name.IndexOf("3D", StringComparison.OrdinalIgnoreCase) >= 0) { PptShape r = sh; sh = null; return r; }
                        }
                        catch { }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return null;
            }
            catch { return null; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private bool GradeProject2Task4()
        {
            Slide slide = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 4);
                if (slide == null) return false;
                PptShape modelShape = null;
                try
                {
                    modelShape = Find3DModelShape(slide);
                    if (modelShape == null) return false;
                    TimeLine timeline = null;
                    try
                    {
                        timeline = slide.TimeLine;
                        if (timeline == null) return false;
                        Sequence seq = null;
                        try
                        {
                            seq = timeline.MainSequence;
                            if (seq == null) return false;
                            int count = seq.Count;
                            for (int i = 1; i <= count; i++)
                            {
                                Effect eff = null;
                                try
                                {
                                    eff = seq[i];
                                    if (eff == null) continue;
                                    try
                                    {
                                        PptShape effShape = eff.Shape;
                                        if (effShape != null)
                                        {
                                            try
                                            {
                                                if (effShape.Id == modelShape.Id)
                                                {
                                                    var et = eff.EffectType;
                                                    try
                                                    {
                                                        // ターンテーブル: msoAnimEffectTurntable または数値 (環境により要確認)
                                                    if (et == (MsoAnimEffect)129) return true;
                                                    }
                                                    catch { }
                                                    try { Marshal.ReleaseComObject(effShape); } catch { }
                                                    break;
                                                }
                                            }
                                            finally { if (effShape != null) try { Marshal.ReleaseComObject(effShape); } catch { } }
                                        }
                                    }
                                    catch { }
                                }
                                finally
                                {
                                    if (eff != null) { try { Marshal.ReleaseComObject(eff); } catch { } }
                                }
                            }
                            return false;
                        }
                        finally
                        {
                            if (seq != null) { try { Marshal.ReleaseComObject(seq); } catch { } }
                        }
                    }
                    finally
                    {
                        if (timeline != null) { try { Marshal.ReleaseComObject(timeline); } catch { } }
                    }
                }
                finally
                {
                    if (modelShape != null) { try { Marshal.ReleaseComObject(modelShape); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        // ----- Project 3: SmartArt・ズーム -----

        private static PptShape FindSmartArtShape(Slide slide)
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
                        if (sh.HasSmartArt == MsoTriState.msoTrue) { PptShape r = sh; sh = null; return r; }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return null;
            }
            catch { return null; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private bool GradeProject3Task1()
        {
            Slide slide = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 5);
                if (slide == null) return false;
                PptShape saShape = null;
                try
                {
                    saShape = FindSmartArtShape(slide);
                    if (saShape == null) return false;
                    SmartArt smartArt = null;
                    try
                    {
                        smartArt = saShape.SmartArt;
                        if (smartArt == null) return false;
                        SmartArtNodes nodes = null;
                        try
                        {
                            nodes = smartArt.AllNodes;
                            if (nodes == null) return false;
                            var sb = new System.Text.StringBuilder();
                            int nCount = nodes.Count;
                            for (int j = 1; j <= nCount; j++)
                            {
                                SmartArtNode node = null;
                                try
                                {
                                    node = nodes[j];
                                    if (node != null)
                                    {
                                        try
                                        {
                                            var tf2 = node.TextFrame2;
                                            if (tf2 != null && tf2.TextRange != null)
                                            {
                                                string t = tf2.TextRange.Text ?? "";
                                                sb.Append(t);
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                finally
                                {
                                    if (node != null) { try { Marshal.ReleaseComObject(node); } catch { } }
                                }
                            }
                            string allText = sb.ToString();
                            return allText.IndexOf("1F受付", StringComparison.OrdinalIgnoreCase) >= 0
                                && allText.IndexOf("面接室", StringComparison.OrdinalIgnoreCase) >= 0;
                        }
                        finally
                        {
                            if (nodes != null) { try { Marshal.ReleaseComObject(nodes); } catch { } }
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
                }
            }
            catch { return false; }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        private bool GradeProject3Task2()
        {
            Slide slide = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 5);
                if (slide == null) return false;
                PptShape saShape = null;
                try
                {
                    saShape = FindSmartArtShape(slide);
                    if (saShape == null) return false;
                    SmartArt smartArt = null;
                    try
                    {
                        smartArt = saShape.SmartArt;
                        if (smartArt == null) return false;
                        try
                        {
                            // 塗りつぶし アクセント5: SmartArt.Color と Application.SmartArtColors で比較
                            var appliedColor = smartArt.Color;
                            if (appliedColor == null) return false;
                            if (_app == null) return false;
                            try
                            {
                                var accent5 = _app.SmartArtColors[14]; // インデックスは環境により要確認
                                if (accent5 != null && appliedColor.Equals(accent5)) return true;
                            }
                            catch { }
                            for (int idx = 1; idx <= 20; idx++)
                            {
                                try
                                {
                                    var style = _app.SmartArtColors[idx];
                                    if (style != null && appliedColor.Equals(style)) return true;
                                }
                                catch { break; }
                            }
                            return false;
                        }
                        catch { return false; }
                    }
                    finally
                    {
                        if (smartArt != null) { try { Marshal.ReleaseComObject(smartArt); } catch { } }
                    }
                }
                finally
                {
                    if (saShape != null) { try { Marshal.ReleaseComObject(saShape); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        // ----- Project 4: 図の書式設定・配置 -----

        private const double PositionTolerance = 2.0;

        private static bool IsPictureShape(PptShape sh)
        {
            try
            {
                int t = (int)sh.Type;
                return t == (int)MsoShapeType.msoPicture || t == 11;
            }
            catch { return false; }
        }

        private bool GradeProject4Task4()
        {
            Slide slide = null;
            try
            {
                slide = GetSlideByTitle(_activePresentation, "THANK YOU");
                if (slide == null) return false;
                PptShapes shapes = null;
                try
                {
                    shapes = slide.Shapes;
                    if (shapes == null) return false;
                    PptShape rightmostPicture = null;
                    float maxRight = float.MinValue;
                    int count = shapes.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        PptShape sh = null;
                        try
                        {
                            sh = shapes[i];
                            if (!IsPictureShape(sh)) continue;
                            float left = (float)sh.Left;
                            float width = (float)sh.Width;
                            float right = left + width;
                            if (right > maxRight)
                            {
                                maxRight = right;
                                if (rightmostPicture != null) { try { Marshal.ReleaseComObject(rightmostPicture); } catch { } }
                                rightmostPicture = sh;
                                sh = null;
                            }
                        }
                        finally
                        {
                            if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                        }
                    }
                    if (rightmostPicture == null) return false;
                    try
                    {
                        Microsoft.Office.Interop.PowerPoint.PictureFormat pf = null;
                        try
                        {
                            pf = (Microsoft.Office.Interop.PowerPoint.PictureFormat)rightmostPicture.PictureFormat;
                            if (pf == null) return false;
                            float cr = pf.CropRight;
                            float cl = pf.CropLeft;
                            float ct = pf.CropTop;
                            float cb = pf.CropBottom;
                            return cr > 0 || cl > 0 || ct > 0 || cb > 0;
                        }
                        finally
                        {
                            if (pf != null) { try { Marshal.ReleaseComObject(pf); } catch { } }
                        }
                    }
                    finally
                    {
                        if (rightmostPicture != null) { try { Marshal.ReleaseComObject(rightmostPicture); } catch { } }
                    }
                }
                finally
                {
                    if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        private bool GradeProject4Task5()
        {
            Slide slide = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 5);
                if (slide == null) return false;
                PptShapes shapes = null;
                try
                {
                    shapes = slide.Shapes;
                    if (shapes == null) return false;
                    PptShape leftPic = null;
                    PptShape rightPic = null;
                    float minLeft = float.MaxValue;
                    float maxLeft = float.MinValue;
                    int count = shapes.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        PptShape sh = null;
                        try
                        {
                            sh = shapes[i];
                            if (!IsPictureShape(sh)) continue;
                            float left = (float)sh.Left;
                            if (left < minLeft)
                            {
                                minLeft = left;
                                if (leftPic != null) { try { Marshal.ReleaseComObject(leftPic); } catch { } }
                                leftPic = sh;
                                sh = null;
                            }
                        }
                        finally
                        {
                            if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                        }
                    }
                    for (int i = 1; i <= count; i++)
                    {
                        PptShape sh = null;
                        try
                        {
                            sh = shapes[i];
                            if (!IsPictureShape(sh)) continue;
                            float left = (float)sh.Left;
                            if (left > maxLeft)
                            {
                                maxLeft = left;
                                if (rightPic != null) { try { Marshal.ReleaseComObject(rightPic); } catch { } }
                                rightPic = sh;
                                sh = null;
                            }
                        }
                        finally
                        {
                            if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                        }
                    }
                    if (leftPic == null || rightPic == null || leftPic.Id == rightPic.Id) return false;
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
                    if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        private bool GradeProject4Task6()
        {
            Slide slide = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 3);
                if (slide == null) return false;
                PptShapes shapes = null;
                try
                {
                    shapes = slide.Shapes;
                    if (shapes == null) return false;
                    PptShape shSmart = null, shTablet = null, shMonitor = null;
                    int count = shapes.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        PptShape sh = null;
                        try
                        {
                            sh = shapes[i];
                            string name = null;
                            string alt = null;
                            try { name = sh.Name ?? ""; } catch { }
                            try { alt = sh.AlternativeText ?? ""; } catch { }
                            string combined = name + " " + alt;
                            if (combined.IndexOf("スマホ", StringComparison.OrdinalIgnoreCase) >= 0) { shSmart = sh; sh = null; }
                            else if (combined.IndexOf("タブレット", StringComparison.OrdinalIgnoreCase) >= 0) { shTablet = sh; sh = null; }
                            else if (combined.IndexOf("モニター", StringComparison.OrdinalIgnoreCase) >= 0) { shMonitor = sh; sh = null; }
                        }
                        finally
                        {
                            if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                        }
                    }
                    if (shSmart == null || shTablet == null || shMonitor == null) return false;
                    try
                    {
                        int zSmart = shSmart.ZOrderPosition;
                        int zTablet = shTablet.ZOrderPosition;
                        int zMonitor = shMonitor.ZOrderPosition;
                        return zSmart > zTablet && zTablet > zMonitor;
                    }
                    finally
                    {
                        if (shSmart != null) { try { Marshal.ReleaseComObject(shSmart); } catch { } }
                        if (shTablet != null) { try { Marshal.ReleaseComObject(shTablet); } catch { } }
                        if (shMonitor != null) { try { Marshal.ReleaseComObject(shMonitor); } catch { } }
                    }
                }
                finally
                {
                    if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        // ----- Project 9: グラフ・ハイパーリンク -----

        private const int XlBarClustered = 57;

        private bool GradeProject9Task1()
        {
            Slide slide = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 2);
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
                            if (sh.HasChart != MsoTriState.msoTrue) continue;
                            Chart chart = null;
                            try
                            {
                                chart = sh.Chart;
                                if (chart == null) continue;
                                try
                                {
                                    int ct = (int)chart.ChartType;
                                    return ct == XlBarClustered;
                                }
                                finally
                                {
                                    if (chart != null) { try { Marshal.ReleaseComObject(chart); } catch { } }
                                }
                            }
                            catch { continue; }
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
            catch { return false; }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        private const string ExpectedHyperlinkAddressTask9_6 = "https://www.jica.go.jp/activities/issues/natural_env/index.html";

        private bool GradeProject9Task6()
        {
            Slide slide = null;
            try
            {
                slide = GetSlideByNumber(_activePresentation, 1);
                if (slide == null) return false;
                PptShape shape = null;
                try
                {
                    shape = FindShapeWithText(slide, "お問い合わせ");
                    if (shape == null) return false;
                    try
                    {
                        Microsoft.Office.Interop.PowerPoint.TextFrame tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)shape.TextFrame;
                        if (tf == null) return false;
                        TextRange tr = null;
                        try
                        {
                            tr = tf.TextRange;
                            if (tr == null) return false;
                            ActionSettings acts = null;
                            try
                            {
                                acts = tr.ActionSettings;
                                if (acts == null) return false;
                                ActionSetting act = null;
                                try
                                {
                                    act = acts[PpMouseActivation.ppMouseClick];
                                    if (act == null) return false;
                                    if (act.Action != PpActionType.ppActionHyperlink) return false;
                                    Hyperlink hyp = null;
                                    try
                                    {
                                        hyp = act.Hyperlink;
                                        if (hyp != null)
                                        {
                                            string addr = (hyp.Address ?? "").Trim();
                                            if (string.Equals(addr, ExpectedHyperlinkAddressTask9_6, StringComparison.OrdinalIgnoreCase))
                                                return true;
                                        }
                                    }
                                    finally
                                    {
                                        if (hyp != null) { try { Marshal.ReleaseComObject(hyp); } catch { } }
                                    }
                                }
                                finally
                                {
                                    if (act != null) { try { Marshal.ReleaseComObject(act); } catch { } }
                                }
                            }
                            finally
                            {
                                if (acts != null) { try { Marshal.ReleaseComObject(acts); } catch { } }
                            }
                            try
                            {
                                int r = 1;
                                while (true)
                                {
                                    TextRange run = null;
                                    try
                                    {
                                        run = tr.Runs(r, 1);
                                        if (run == null) break;
                                        ActionSettings runActs = null;
                                        try
                                        {
                                            runActs = run.ActionSettings;
                                            if (runActs == null) { r++; continue; }
                                            ActionSetting runAct = null;
                                            try
                                            {
                                                runAct = runActs[PpMouseActivation.ppMouseClick];
                                                if (runAct != null && runAct.Action == PpActionType.ppActionHyperlink)
                                                {
                                                    Hyperlink runHyp = null;
                                                    try
                                                    {
                                                        runHyp = runAct.Hyperlink;
                                                        if (runHyp != null)
                                                        {
                                                            string addr = (runHyp.Address ?? "").Trim();
                                                            if (string.Equals(addr, ExpectedHyperlinkAddressTask9_6, StringComparison.OrdinalIgnoreCase))
                                                                return true;
                                                        }
                                                    }
                                                    finally
                                                    {
                                                        if (runHyp != null) { try { Marshal.ReleaseComObject(runHyp); } catch { } }
                                                    }
                                                }
                                            }
                                            finally
                                            {
                                                if (runAct != null) { try { Marshal.ReleaseComObject(runAct); } catch { } }
                                            }
                                        }
                                        finally
                                        {
                                            if (runActs != null) { try { Marshal.ReleaseComObject(runActs); } catch { } }
                                        }
                                    }
                                    finally
                                    {
                                        if (run != null) { try { Marshal.ReleaseComObject(run); } catch { } }
                                    }
                                    r++;
                                }
                            }
                            catch { }
                            return false;
                        }
                        finally
                        {
                            if (tr != null) { try { Marshal.ReleaseComObject(tr); } catch { } }
                        }
                    }
                    finally
                    {
                        if (shape != null) { try { Marshal.ReleaseComObject(shape); } catch { } }
                    }
                }
                catch { return false; }
            }
            finally
            {
                if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (_disposed)
                return;
            if (disposing)
            {
                try
                {
                    if (_activePresentation != null)
                    {
                        Marshal.ReleaseComObject(_activePresentation);
                        _activePresentation = null;
                    }
                }
                catch { }
                try
                {
                    if (_app != null)
                    {
                        Marshal.ReleaseComObject(_app);
                        _app = null;
                    }
                }
                catch { }
            }
            _disposed = true;
        }
    }
}
