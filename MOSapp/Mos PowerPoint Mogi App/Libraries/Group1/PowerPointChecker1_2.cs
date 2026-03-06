using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_2
    {
        /// <summary>2-1: すべてのスライドに画面切り替え「プッシュ・右から」。2-3で3,4,5は渦巻きになるため、1,2,6のみチェック。3854(右から)と3853(左から)を許容（日本語UIで1枚目が3853になる環境あり）。</summary>
        public bool CheckTask_1_2_01()
        {
            const int ppEffectPushRight = 3854;
            const int ppEffectPushLeft = 3853;
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
                                if (effectVal != ppEffectPushRight && effectVal != ppEffectPushLeft) return false;
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

        /// <summary>2-2: すべての画面切り替えを3秒に設定。</summary>
        public bool CheckTask_1_2_02()
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
                                if (isSlide345)
                                {
                                    if ((dur < 2.9f || dur > 3.1f) && (dur < 3.9f || dur > 4.1f)) return false;
                                }
                                else
                                {
                                    if (dur < 2.9f || dur > 3.1f) return false;
                                }
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

        /// <summary>2-3: スライド3,4,5に「渦巻き」の画面切り替えを設定。</summary>
        public bool CheckTask_1_2_03()
        {
            const int ppEffectSpiral = 3357;
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
                            bool ok = (effectVal == ppEffectSpiral) || (effectVal >= 3863 && effectVal <= 3866);
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
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>2-4: スライド4の3Dモデルに「ターンテーブル」のアニメーションを設定。</summary>
        public bool CheckTask_1_2_04()
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
                    PptShape modelShape = null;
                    try
                    {
                        modelShape = PowerPointCheckerCommon.Find3DModelShape(slide);
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
                                                        int etVal = (int)eff.EffectType;
                                                        if (etVal == 129 || etVal == 152) return true;
                                                        try
                                                        {
                                                            string dn = eff.DisplayName ?? "";
                                                            if (dn.IndexOf("turn", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                                                dn.IndexOf("ターン", StringComparison.OrdinalIgnoreCase) >= 0)
                                                                return true;
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
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>2-5: スライド2の人型画像のアニメーションをワイプアウト（横）、継続時間2秒に設定。</summary>
        public bool CheckTask_1_2_05()
        {
            const int msoAnimEffectWipe = 22;
            const int msoAnimEffectSplit = 16;
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
                        int sc = shapes.Count;
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
                                for (int i = 1; i <= sc; i++)
                                {
                                    PptShape sh = null;
                                    try
                                    {
                                        sh = shapes[i];
                                        bool isPic = PowerPointCheckerCommon.IsPictureShape(sh);
                                        if (!isPic) continue;
                                        int shapeId = sh.Id;
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
                                                        if (es.Id != shapeId) continue;
                                                        int etVal = (int)eff.EffectType;
                                                        if (etVal != msoAnimEffectWipe && etVal != msoAnimEffectSplit) continue;
                                                        try
                                                        {
                                                            float dur = (float)eff.Timing.Duration;
                                                            if (dur >= 1.9f && dur <= 2.1f) return true;
                                                        }
                                                        catch { }
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

        /// <summary>2-6: スライド「今年度募集について」の箇条書きアニメーションをひし形・段落別に設定。</summary>
        public bool CheckTask_1_2_06()
        {
            const int msoAnimEffectDiamond = 8;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, "今年度募集について");
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
                                    if (eff == null) continue;
                                    try
                                    {
                                        var et = eff.EffectType;
                                        if ((int)et != msoAnimEffectDiamond) continue;
                                        try
                                        {
                                            var info = eff.EffectInformation;
                                            if (info != null && info.TextUnitEffect == MsoAnimTextUnitEffect.msoAnimTextUnitEffectByParagraph)
                                                return true;
                                        }
                                        catch { }
                                        try
                                        {
                                            if (eff.Paragraph != 0) return true;
                                        }
                                        catch { }
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
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>2-7: スライド6の星の画像にアニメーションの軌跡「プラス」を設定。</summary>
        public bool CheckTask_1_2_07()
        {
            const int msoAnimEffectPathPlus = 117;
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
                                    if (eff == null) continue;
                                    try
                                    {
                                        var et = eff.EffectType;
                                        if ((int)et == msoAnimEffectPathPlus) return true;
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
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
    }
}
