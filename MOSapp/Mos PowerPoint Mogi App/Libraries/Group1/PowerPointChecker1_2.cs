using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace Libraries.Group1
{
    public class PowerPointChecker1_2
    {
        public bool CheckTask_1_2_01()
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
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_2_02() { return false; }
        public bool CheckTask_1_2_03() { return false; }

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
                                                        var et = eff.EffectType;
                                                        try
                                                        {
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
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_2_05() { return false; }
        public bool CheckTask_1_2_06() { return false; }
        public bool CheckTask_1_2_07() { return false; }
    }
}
