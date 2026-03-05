using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using Libraries;

namespace Libraries.Group1
{
    public class PowerPointChecker1_8
    {
        /// <summary>8-1: スライド「スクールの様子」に動画が挿入されているか。</summary>
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
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, "スクールの様子");
                    if (slide == null) return false;
                    PptShape videoShape = null;
                    try
                    {
                        videoShape = PowerPointCheckerCommon.FindFirstVideoShape(slide);
                        return videoShape != null;
                    }
                    finally
                    {
                        if (videoShape != null) { try { Marshal.ReleaseComObject(videoShape); } catch { } }
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

        /// <summary>8-2: スライド5にビデオが挿入されているか。</summary>
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
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    if (slide == null) return false;
                    PptShape videoShape = null;
                    try
                    {
                        videoShape = PowerPointCheckerCommon.FindFirstVideoShape(slide);
                        return videoShape != null;
                    }
                    finally
                    {
                        if (videoShape != null) { try { Marshal.ReleaseComObject(videoShape); } catch { } }
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

        /// <summary>8-3: スライド5のビデオを開始00:05・終了00:10にトリム。</summary>
        public bool CheckTask_1_8_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    if (slide == null) return false;
                    PptShape videoShape = null;
                    try
                    {
                        videoShape = PowerPointCheckerCommon.FindFirstVideoShape(slide);
                        if (videoShape == null) return false;
                        MediaFormat mf = null;
                        try
                        {
                            mf = videoShape.MediaFormat;
                            if (mf == null) return false;
                            try
                            {
                                float startPt = (float)mf.StartPoint;
                                float endPt = (float)mf.EndPoint;
                                return Math.Abs(startPt - 5000f) < 500f && Math.Abs(endPt - 10000f) < 500f;
                            }
                            catch { return false; }
                        }
                        finally { if (mf != null) { try { Marshal.ReleaseComObject(mf); } catch { } } }
                    }
                    finally { if (videoShape != null) { try { Marshal.ReleaseComObject(videoShape); } catch { } } }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>8-4: スライド1のオーディオをスライド切り替えでも1回再生・フェードイン4秒・繰り返し。COM の FadeInDuration または VSTO ログで判定。</summary>
        public bool CheckTask_1_8_04()
        {
            if (PPLogReader.HasTask8_4AudioExecuted())
                return true;
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
                                try
                                {
                                    if (sh.MediaType != PpMediaType.ppMediaTypeSound) continue;
                                }
                                catch { continue; }
                                MediaFormat mf = null;
                                try
                                {
                                    mf = sh.MediaFormat;
                                    if (mf == null) continue;
                                    try
                                    {
                                        // PlayAcrossSlides/RewindAfterPlaying は PIA で未定義。VSTO で検証予定。
                                        float fadeIn = (float)mf.FadeInDuration;
                                        return Math.Abs(fadeIn - 4000f) < 500f;
                                    }
                                    catch { return false; }
                                }
                                finally { if (mf != null) { try { Marshal.ReleaseComObject(mf); } catch { } } }
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

        /// <summary>8-5: プレゼンテーションを読み取り専用に設定。</summary>
        public bool CheckTask_1_8_05()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                try
                {
                    return pres.ReadOnly == MsoTriState.msoTrue;
                }
                catch { return false; }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
    }
}
