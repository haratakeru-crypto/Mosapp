using System;
using System.Runtime.InteropServices;
using System.Threading;
using Libraries;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    /// <summary>プロジェクト9（P9-1〜P9-5）。Phase B でタスク単位に Legacy 移植。</summary>
    public class PowerPointChecker1_9
    {
        private const string ExpectedSlideTitleP9_1 = "スクールの様子";
        private const int P9_2TargetSlideNumber = 5;
        private const float P9_2ExpectedStartMs = 4000f;
        private const float P9_2ExpectedEndMs = 9000f;
        private const float P9_2TrimToleranceMs = 500f;
        private const int P9_3TargetSlideNumber = 1;
        private const string P9_3PlayAcrossMarker = "[Task9-3] PlayAcrossSlides";
        private const string P9_3FadeOutMarker = "[Task9-3] FadeOut3000";
        private const int P9_4TargetSlideNumber = 2;
        /// <summary>Excel XlChartType.xlColumnClustered（集合縦棒）。</summary>
        private const int XlColumnClustered = 51;
        private const int P9_5TargetSlideNumber = 2;

        /// <summary>P9-1: スライド「スクールの様子」に動画挿入（旧8-1）。アイコン表示は COM では区別せず動画存在で判定。</summary>
        public bool CheckTask_1_9_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, ExpectedSlideTitleP9_1);
                    if (slide == null) return false;
                    PptShape videoShape = null;
                    try
                    {
                        if (!TryFindVideoForScoring(slide, out videoShape))
                            return false;
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
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        /// <summary>P9-2: スライド5のビデオを開始4秒・終了9秒にトリム（旧8-3、5秒/10秒→4秒/9秒）。</summary>
        public bool CheckTask_1_9_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, P9_2TargetSlideNumber);
                    if (slide == null) return false;
                    PptShape videoShape = null;
                    try
                    {
                        if (!TryFindVideoForScoring(slide, out videoShape) || videoShape == null)
                            return false;

                        MediaFormat mf = null;
                        try
                        {
                            mf = videoShape.MediaFormat;
                            if (mf == null) return false;
                            float startPt = (float)mf.StartPoint;
                            float endPt = (float)mf.EndPoint;
                            return Math.Abs(startPt - P9_2ExpectedStartMs) < P9_2TrimToleranceMs
                                && Math.Abs(endPt - P9_2ExpectedEndMs) < P9_2TrimToleranceMs;
                        }
                        catch { return false; }
                        finally
                        {
                            if (mf != null) { try { Marshal.ReleaseComObject(mf); } catch { } }
                        }
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
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        /// <summary>P9-3: スライド1オーディオ・スライド切替後も再生＋フェードアウト3秒（旧8-4）。両条件 AND。</summary>
        public bool CheckTask_1_9_03()
        {
            if (TrySlide1AudioP9_3ViaCom())
                return true;

            bool playAcross = PPLogReader.HasMarkerWithinTask(9, 3, P9_3PlayAcrossMarker);
            if (!playAcross)
                return false;
            return PPLogReader.HasMarkerWithinTask(9, 3, P9_3FadeOutMarker);
        }

        private static bool TrySlide1AudioP9_3ViaCom()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                bool batch = PPLogReader.IsBatchScoring();
                int maxAttempts = batch ? 2 : 5;
                for (int attempt = 1; attempt <= maxAttempts; attempt++)
                {
                    Slide slide = null;
                    try
                    {
                        slide = PowerPointCheckerCommon.GetSlideByNumber(pres, P9_3TargetSlideNumber);
                        if (slide == null)
                        {
                            if (batch) return false;
                            continue;
                        }

                        PowerPointCheckerCommon.MediaScanStatus status = ScanSlideAudioP9_3(slide, pres);
                        if (status == PowerPointCheckerCommon.MediaScanStatus.Found)
                            return true;
                        if (batch && status == PowerPointCheckerCommon.MediaScanStatus.NotFound)
                            return false;
                    }
                    finally
                    {
                        if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                    }
                    if (attempt < maxAttempts)
                        Thread.Sleep(batch ? 100 : 400);
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        private static bool TryFindVideoForScoring(Slide slide, out PptShape videoShape)
        {
            videoShape = null;
            bool batch = PPLogReader.IsBatchScoring();
            int maxAttempts = batch ? 2 : 5;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                PowerPointCheckerCommon.MediaScanStatus status =
                    PowerPointCheckerCommon.TryFindFirstVideoShape(slide, out videoShape);
                if (status == PowerPointCheckerCommon.MediaScanStatus.Found)
                    return videoShape != null;
                if (videoShape != null)
                {
                    try { Marshal.ReleaseComObject(videoShape); } catch { }
                    videoShape = null;
                }
                if (batch && status == PowerPointCheckerCommon.MediaScanStatus.NotFound)
                    return false;
                if (attempt < maxAttempts)
                    Thread.Sleep(batch ? 100 : 400);
            }
            return videoShape != null;
        }

        private static PowerPointCheckerCommon.MediaScanStatus ScanSlideAudioP9_3(Slide slide, Presentation pres)
        {
            if (slide == null) return PowerPointCheckerCommon.MediaScanStatus.TransientComFailure;
            PptShapes shapes = null;
            bool sawIndexFailure = false;
            try
            {
                try { shapes = slide.Shapes; }
                catch { return PowerPointCheckerCommon.MediaScanStatus.TransientComFailure; }
                if (shapes == null) return PowerPointCheckerCommon.MediaScanStatus.TransientComFailure;
                int shapeCount;
                try { shapeCount = shapes.Count; }
                catch { return PowerPointCheckerCommon.MediaScanStatus.TransientComFailure; }

                for (int i = 1; i <= shapeCount; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        try { sh = shapes[i]; }
                        catch
                        {
                            sawIndexFailure = true;
                            continue;
                        }

                        if (PptAudioMediaHelper.ShapeMatchesP9_3AudioSettings(sh, pres))
                            return PowerPointCheckerCommon.MediaScanStatus.Found;
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return sawIndexFailure
                    ? PowerPointCheckerCommon.MediaScanStatus.TransientComFailure
                    : PowerPointCheckerCommon.MediaScanStatus.NotFound;
            }
            catch { return PowerPointCheckerCommon.MediaScanStatus.TransientComFailure; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        /// <summary>P9-4: スライド2に集合縦棒グラフ（旧9-1）。ChartType = xlColumnClustered のみ。</summary>
        public bool CheckTask_1_9_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, P9_4TargetSlideNumber);
                    if (slide == null) return false;
                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
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
                                if (sh.HasChart != MsoTriState.msoTrue) continue;
                                Chart chart = null;
                                try
                                {
                                    chart = sh.Chart;
                                    if (chart == null) continue;
                                    int ct;
                                    try { ct = (int)chart.ChartType; }
                                    catch { continue; }
                                    if (ct == XlColumnClustered)
                                        return true;
                                }
                                catch { continue; }
                                finally
                                {
                                    if (chart != null) { try { Marshal.ReleaseComObject(chart); } catch { } }
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
                        if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
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

        /// <summary>P9-5: スライド2グラフにデータテーブル（凡例マーカーなし）・タイトル/凡例削除（旧9-3）。</summary>
        public bool CheckTask_1_9_05()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, P9_5TargetSlideNumber);
                    if (slide == null) return false;
                    return SlideHasChartMatchingP9_5(slide);
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

        private static bool SlideHasChartMatchingP9_5(Slide slide)
        {
            if (slide == null) return false;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
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
                        if (sh.HasChart != MsoTriState.msoTrue) continue;
                        Chart chart = null;
                        try
                        {
                            chart = sh.Chart;
                            if (chart != null && IsChartMatchingP9_5(chart))
                                return true;
                        }
                        catch { continue; }
                        finally
                        {
                            if (chart != null) { try { Marshal.ReleaseComObject(chart); } catch { } }
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

        private static bool IsChartMatchingP9_5(Chart chart)
        {
            if (chart == null) return false;
            try
            {
                if (!chart.HasDataTable) return false;
                if (chart.HasTitle) return false;
                if (chart.HasLegend) return false;

                object dataTable = null;
                try
                {
                    dataTable = chart.DataTable;
                    if (dataTable == null) return false;
                    dynamic dt = dataTable;
                    return !(bool)dt.ShowLegendKey;
                }
                catch { return false; }
                finally
                {
                    if (dataTable != null) { try { Marshal.ReleaseComObject(dataTable); } catch { } }
                }
            }
            catch { return false; }
        }
    }
}
