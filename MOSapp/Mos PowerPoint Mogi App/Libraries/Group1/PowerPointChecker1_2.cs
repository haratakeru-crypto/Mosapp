using System;
using System.Runtime.InteropServices;
using Libraries;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_2
    {
        /// <summary>
        /// P2-1: 画面切り替え「スプリット」＋「ワイプアウト（横）」(ppEffectSplitHorizontalOut=3585)。
        /// 問題文は全スライドだが、P2-3 で 3〜5 が上書きされるため採点は 1,2,6 のみ（6枚構成時）。
        /// P2-4 で上書きされた場合は VSTO 証跡でフォールバック。
        /// </summary>
        public bool CheckTask_1_2_01()
        {
            if (PPLogReader.HasTask2_1SplitHorizontalOutExecuted())
                return true;
            const int ppEffectSplitHorizontalOut = 3585;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null) return false;
                    int count = slides.Count;
                    if (count < 2) return false;
                    int[] indicesToCheck = count >= 6 ? new[] { 1, 2, 6 } : new[] { 1, 2 };
                    foreach (int i in indicesToCheck)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            try
                            {
                                int effectVal = (int)slide.SlideShowTransition.EntryEffect;
                                if (effectVal != ppEffectSplitHorizontalOut) return false;
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
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>
        /// P2-2: すべての画面切り替えの継続時間を3秒に設定。
        /// スライド3〜5は P2-3 の「切り替え」適用後に約1.25秒になるため、3秒または1.25秒を許容。
        /// P2-4 で上書きされた場合は VSTO 証跡でフォールバック。
        /// </summary>
        public bool CheckTask_1_2_02()
        {
            if (PPLogReader.HasTask2_2TransitionDuration3SecExecuted())
                return true;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null) return false;
                    int count = slides.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            try
                            {
                                float dur = slide.SlideShowTransition.Duration;
                                bool isSlide345 = (i == 3 || i == 4 || i == 5);
                                bool ok = isSlide345
                                    ? IsTransitionDurationThreeSeconds(dur) || IsTransitionDurationSwitchDefault(dur)
                                    : IsTransitionDurationThreeSeconds(dur);
                                if (!ok) return false;
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
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>
        /// P2-3: スライド3,4,5に「切り替え」(ppEffectSwitchRight=3903)の画面切り替えを設定。
        /// P2-4 で上書きされた場合は VSTO 証跡でフォールバック。
        /// </summary>
        public bool CheckTask_1_2_03()
        {
            if (PPLogReader.HasTask2_3SwitchRightExecuted())
                return true;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                for (int slideNum = 3; slideNum <= 5; slideNum++)
                {
                    Slide slide = null;
                    try
                    {
                        slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
                        if (slide == null) return false;
                        try
                        {
                            int effectVal = (int)slide.SlideShowTransition.EntryEffect;
                            if (!IsP2SwitchTransition(effectVal)) return false;
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
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>
        /// P2-4: すべてのスライドが 5 秒後に自動で次のスライドへ進む
        /// （画面切り替えの [自動] が ON かつ [5] 秒）
        /// </summary>
        public bool CheckTask_1_2_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null) return false;

                    int count = slides.Count;
                    if (count <= 0) return false;

                    for (int i = 1; i <= count; i++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            if (slide == null) return false;

                            SlideShowTransition transition = null;
                            try
                            {
                                transition = slide.SlideShowTransition;
                                if (transition == null) return false;

                                bool advanceOnTime;
                                float advanceTime;
                                try
                                {
                                    advanceOnTime = transition.AdvanceOnTime == MsoTriState.msoTrue;
                                    advanceTime = transition.AdvanceTime;
                                }
                                catch
                                {
                                    return false;
                                }

                                if (!advanceOnTime) return false;

                                // UI 入力誤差や保存差を考慮して 5.0 秒 ±0.1 を許容
                                if (advanceTime < 4.9f || advanceTime > 5.1f) return false;
                            }
                            finally
                            {
                                if (transition != null) { try { Marshal.ReleaseComObject(transition); } catch { } }
                            }
                        }
                        finally
                        {
                            if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                        }
                    }

                    return true;
                }
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        /// <summary>
        /// P2-5: スライド6の3Dモデル「黒板」にアニメーション「ジャンプしてターン」(EffectType=154)を設定。
        /// 3Dモデルは DisplayName が図形名になるため EffectType で判定する。
        /// </summary>
        public bool CheckTask_1_2_05()
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
                    PptShape modelShape = null;
                    try
                    {
                        modelShape = FindP2Blackboard3DModel(slide);
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
                                int modelId = modelShape.Id;
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
                                                    if (effShape.Id == modelId && IsP2JumpTurnAnimation(eff))
                                                        return true;
                                                }
                                                finally { try { Marshal.ReleaseComObject(effShape); } catch { } }
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
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>
        /// P2-6: スライド2の2つの画像が左上隅から飛び込み、継続時間0.55秒。
        /// 対象画像が2枚以上アニメーション条件を満たせば合格。
        /// </summary>
        public bool CheckTask_1_2_06()
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
                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
                        if (shapes == null) return false;
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
                                int matched = 0;
                                int sc = shapes.Count;
                                for (int i = 1; i <= sc; i++)
                                {
                                    PptShape sh = null;
                                    try
                                    {
                                        sh = shapes[i];
                                        if (!PowerPointCheckerCommon.IsPictureShape(sh)) continue;
                                        if (ShapeHasP2FlyFromTopLeftAnimation(seq, sh.Id))
                                            matched++;
                                    }
                                    finally
                                    {
                                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                                    }
                                }
                                return matched >= 2;
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
                        if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
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

        /// <summary>
        /// P2-7: スライド6の箇条書きアニメーションを「プラス」にし、効果のオプション「すべて同時」に設定。
        /// </summary>
        public bool CheckTask_1_2_07()
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
                                    if (eff == null || !IsP2PlusAnimation(eff)) continue;
                                    PptShape es = null;
                                    try
                                    {
                                        es = eff.Shape;
                                        if (es == null || !ShapeHasBulletParagraph(es)) continue;
                                        if (IsP2PlusAllAtOnceEffect(eff))
                                            return true;
                                    }
                                    finally
                                    {
                                        if (es != null) { try { Marshal.ReleaseComObject(es); } catch { } }
                                    }
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
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>
        /// P2-8: スライド4の星に軌跡「ニュートロン」、円にはアニメーションなし。
        /// </summary>
        public bool CheckTask_1_2_08()
        {
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
                    PptShape starShape = null;
                    PptShape circleShape = null;
                    try
                    {
                        starShape = FindP2StarShape(slide);
                        circleShape = FindP2CircleShape(slide);
                        if (starShape == null || circleShape == null) return false;
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
                                if (!ShapeHasP2NeutronPathAnimation(seq, starShape.Id)) return false;
                                if (ShapeHasAnimationInSequence(seq, circleShape.Id)) return false;
                                return true;
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
                        if (starShape != null) { try { Marshal.ReleaseComObject(starShape); } catch { } }
                        if (circleShape != null) { try { Marshal.ReleaseComObject(circleShape); } catch { } }
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

        private static bool IsTransitionDurationThreeSeconds(float duration)
        {
            return duration >= 2.9f && duration <= 3.1f;
        }

        /// <summary>P2-3「切り替え」適用後の既定継続時間（約1.25秒）。</summary>
        private static bool IsTransitionDurationSwitchDefault(float duration)
        {
            return duration >= 1.15f && duration <= 1.35f;
        }

        /// <summary>日本語UI「切り替え」の既定オプション = ppEffectSwitchRight (3903)。</summary>
        private static bool IsP2SwitchTransition(int entryEffect)
        {
            const int ppEffectSwitchRight = 3903;
            return entryEffect == ppEffectSwitchRight;
        }

        private static PptShape FindP2Blackboard3DModel(Slide slide)
        {
            PptShape model = PowerPointCheckerCommon.Find3DModelShapeByName(slide, "黒板");
            if (model != null) return model;
            model = PowerPointCheckerCommon.Find3DModelShapeByName(slide, "Blackboard");
            if (model != null) return model;
            return PowerPointCheckerCommon.Find3DModelShape(slide);
        }

        private static bool IsP2JumpTurnAnimation(Effect eff)
        {
            if (eff == null) return false;
            const int pp3DAnimJumpTurn = 154;
            const int pp3DAnimTurntable = 152;
            try
            {
                int etVal = (int)eff.EffectType;
                if (etVal == pp3DAnimTurntable) return false;
                if (etVal == pp3DAnimJumpTurn) return true;
            }
            catch { }
            try
            {
                string dn = eff.DisplayName ?? "";
                if (dn.IndexOf("ターンテーブル", StringComparison.OrdinalIgnoreCase) >= 0) return false;
                if (dn.IndexOf("Turntable", StringComparison.OrdinalIgnoreCase) >= 0) return false;
                if (dn.IndexOf("ジャンプしてターン", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                if (dn.IndexOf("ジャンプ", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                if (dn.IndexOf("Jump", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    dn.IndexOf("Turn", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            catch { }
            return false;
        }

        private static bool ShapeHasP2FlyFromTopLeftAnimation(Sequence seq, int shapeId)
        {
            if (seq == null) return false;
            int seqCount = seq.Count;
            for (int j = 1; j <= seqCount; j++)
            {
                Effect eff = null;
                try
                {
                    eff = seq[j];
                    if (eff == null) continue;
                    try
                    {
                        PptShape es = eff.Shape;
                        if (es == null) continue;
                        try
                        {
                            if (es.Id == shapeId && IsP2FlyFromTopLeftAnimation(eff))
                                return true;
                        }
                        finally { try { Marshal.ReleaseComObject(es); } catch { } }
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

        /// <summary>P2-6 アニメーション継続時間0.55秒（±0.01）。</summary>
        private static bool IsP2FlyAnimationDuration(float duration)
        {
            const float target = 0.55f;
            const float tolerance = 0.01f;
            return duration >= target - tolerance && duration <= target + tolerance;
        }

        private static bool IsP2PlusAnimation(Effect eff)
        {
            if (eff == null) return false;
            const int msoAnimEffectPlus = 13;
            try
            {
                if ((int)eff.EffectType == msoAnimEffectPlus) return true;
            }
            catch { }
            try
            {
                string dn = eff.DisplayName ?? "";
                if (dn.IndexOf("プラス", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                if (dn.Equals("Plus", StringComparison.OrdinalIgnoreCase)) return true;
            }
            catch { }
            return false;
        }

        /// <summary>効果のオプション「すべて同時」(BuildByLevelEffect=AllLevels)。Legacy 2-6 段落別の逆。</summary>
        private static bool IsP2PlusAllAtOnceEffect(Effect eff)
        {
            if (eff == null) return false;
            const int msoAnimateTextByAllLevels = 1;
            EffectInformation info = null;
            try { info = eff.EffectInformation; } catch { }
            if (info != null)
            {
                try
                {
                    int buildByLevel = (int)info.BuildByLevelEffect;
                    if (buildByLevel == msoAnimateTextByAllLevels)
                        return true;
                    if (buildByLevel > msoAnimateTextByAllLevels)
                        return false;
                }
                catch { }
                try
                {
                    if (info.TextUnitEffect == MsoAnimTextUnitEffect.msoAnimTextUnitEffectByParagraph)
                        return false;
                }
                catch { return false; }
            }
            try
            {
                if (eff.Paragraph != 0) return false;
            }
            catch { return false; }
            return true;
        }

        private static PptShape FindP2StarShape(Slide slide)
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
                        if (IsStarAutoShape(sh))
                        {
                            PptShape result = sh;
                            sh = null;
                            return result;
                        }
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

        private static PptShape FindP2CircleShape(Slide slide)
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
                        if (IsCircleAutoShape(sh))
                        {
                            PptShape result = sh;
                            sh = null;
                            return result;
                        }
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

        private static bool IsStarAutoShape(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                if (sh.Type != MsoShapeType.msoAutoShape) return false;
                int t = (int)sh.AutoShapeType;
                const int msoShape8pointStar = 58;
                const int msoShape16pointStar = 59;
                const int msoShape24pointStar = 60;
                const int msoShape32pointStar = 61;
                const int msoShape5pointStar = 92;
                return t == msoShape5pointStar || t == msoShape8pointStar || t == msoShape16pointStar
                    || t == msoShape24pointStar || t == msoShape32pointStar;
            }
            catch { return false; }
        }

        private static bool IsCircleAutoShape(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                if (sh.Type != MsoShapeType.msoAutoShape) return false;
                const int msoShapeOval = 9;
                return (int)sh.AutoShapeType == msoShapeOval;
            }
            catch { return false; }
        }

        private static bool ShapeHasAnimationInSequence(Sequence seq, int shapeId)
        {
            if (seq == null) return false;
            int count = seq.Count;
            for (int i = 1; i <= count; i++)
            {
                Effect eff = null;
                try
                {
                    eff = seq[i];
                    if (eff == null) continue;
                    PptShape es = null;
                    try
                    {
                        es = eff.Shape;
                        if (es != null && es.Id == shapeId) return true;
                    }
                    finally
                    {
                        if (es != null) { try { Marshal.ReleaseComObject(es); } catch { } }
                    }
                }
                finally
                {
                    if (eff != null) { try { Marshal.ReleaseComObject(eff); } catch { } }
                }
            }
            return false;
        }

        private static bool ShapeHasP2NeutronPathAnimation(Sequence seq, int shapeId)
        {
            if (seq == null) return false;
            int count = seq.Count;
            for (int i = 1; i <= count; i++)
            {
                Effect eff = null;
                try
                {
                    eff = seq[i];
                    if (eff == null) continue;
                    PptShape es = null;
                    try
                    {
                        es = eff.Shape;
                        if (es != null && es.Id == shapeId && IsP2NeutronPathAnimation(eff))
                            return true;
                    }
                    finally
                    {
                        if (es != null) { try { Marshal.ReleaseComObject(es); } catch { } }
                    }
                }
                finally
                {
                    if (eff != null) { try { Marshal.ReleaseComObject(eff); } catch { } }
                }
            }
            return false;
        }

        private static bool IsP2NeutronPathAnimation(Effect eff)
        {
            if (eff == null) return false;
            const int msoAnimEffectPathNeutron = 114;
            const int msoAnimEffectPathPlus = 117;
            try
            {
                int etVal = (int)eff.EffectType;
                if (etVal == msoAnimEffectPathPlus) return false;
                if (etVal == msoAnimEffectPathNeutron) return true;
            }
            catch { }
            try
            {
                string dn = eff.DisplayName ?? "";
                if (dn.IndexOf("ニュートロン", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                if (dn.IndexOf("Neutron", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            catch { }
            return false;
        }

        private static bool ShapeHasBulletParagraph(PptShape sh)
        {
            if (sh == null) return false;
            if (sh.HasTextFrame != MsoTriState.msoTrue) return false;
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
            finally
            {
                if (tr != null) { try { Marshal.ReleaseComObject(tr); } catch { } }
            }
        }

        /// <summary>飛び込み（Fly）＋左上方向、継続時間0.55秒。</summary>
        private static bool IsP2FlyFromTopLeftAnimation(Effect eff)
        {
            if (eff == null) return false;
            const int msoAnimEffectFly = 2;
            const int msoAnimDirectionUpLeft = 6;
            const int msoAnimDirectionTopLeft = 12;
            try
            {
                if ((int)eff.EffectType != msoAnimEffectFly) return false;
                float dur = (float)eff.Timing.Duration;
                if (!IsP2FlyAnimationDuration(dur)) return false;
                int dirVal = -1;
                try { dirVal = (int)eff.EffectParameters.Direction; } catch { return false; }
                return dirVal == msoAnimDirectionUpLeft || dirVal == msoAnimDirectionTopLeft;
            }
            catch { return false; }
        }
    }
}
