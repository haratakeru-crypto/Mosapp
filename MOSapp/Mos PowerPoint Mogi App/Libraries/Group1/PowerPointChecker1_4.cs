using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_4
    {
        private const double PositionTolerance = 2.0;

        /// <summary>4-1: スライド1の画像に「四角形 ぼかし」スタイル＋「フィルム粒子」アート効果。</summary>
        public bool CheckTask_1_4_01()
        {
            Presentation pres = null;
            string tempPath = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                // ---- スタイル判定: COMでSoftEdgeを確認 ----
                bool styleOk = false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                    if (slide != null)
                    {
                        PptShapes shapes = slide.Shapes;
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                dynamic dSh = sh;
                                // SoftEdge（ぼかし）チェック
                                try
                                {
                                    dynamic se = dSh.SoftEdge;
                                    if (se != null)
                                    {
                                        if ((int)se.Type >= 1 || (float)se.Radius > 0) styleOk = true;
                                        Marshal.ReleaseComObject(se);
                                    }
                                } catch { }
                                // PictureStyle補完チェック
                                if (!styleOk)
                                {
                                    try
                                    {
                                        dynamic pf = dSh.PictureFormat;
                                        int ps = (int)pf.PictureStyle;
                                        if (ps == 13 || ps == 20 || ps == 21) styleOk = true;
                                        Marshal.ReleaseComObject(pf);
                                    } catch { }
                                }
                            }
                            catch { }
                            finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                            if (styleOk) break;
                        }
                    }
                }
                finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }

                // ---- アート効果判定: pptx XML直接解析 ----
                bool effectOk = false;
                try
                {
                    tempPath = Path.Combine(Path.GetTempPath(), "mos_1_4_01_check_" + Guid.NewGuid().ToString("N") + ".pptx");
                    pres.SaveCopyAs(tempPath);

                    using (var package = Package.Open(tempPath, FileMode.Open, FileAccess.Read))
                    {
                        foreach (var part in package.GetParts())
                        {
                            if (!part.Uri.OriginalString.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;

                            string xml;
                            try
                            {
                                using (var reader = new StreamReader(part.GetStream()))
                                    xml = reader.ReadToEnd();
                            }
                            catch { continue; }

                            // フィルム粒子候補: artisticGrain / artisticFilmGrain / FilmGrain
                            if (Regex.IsMatch(xml, @"artistic(Grain|FilmGrain|Film)", RegexOptions.IgnoreCase)
                                || Regex.IsMatch(xml, @"FilmGrain", RegexOptions.IgnoreCase))
                            {
                                effectOk = true;
                                break;
                            }
                        }
                    }
                }
                catch { }

                return styleOk && effectOk;
            }
            catch { return false; }
            finally
            {
                if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { }
                if (tempPath != null && File.Exists(tempPath))
                    try { File.Delete(tempPath); } catch { }
            }
        }

        /// <summary>4-2: スライド1の画像の図の効果を「反射（中）、オフセットなし」。</summary>
        public bool CheckTask_1_4_02()
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
                    PptShape picture = GetFirstPictureOnSlide(slide);
                    if (picture == null) return false;
                    try
                    {
                        dynamic reflObj = picture.Reflection;
                        if (reflObj == null) return false;
                        try
                        {
                            int t = (int)reflObj.Type;
                            double offset = 0;
                            try { offset = (double)reflObj.Offset; } catch { }
                            
                            bool typeMedium = (t == 2 || t == 5); 
                            bool noOffset = Math.Abs(offset) < 0.05;

                            return typeMedium && noOffset;
                        }
                        catch
                        {
                            try
                            {
                                int t = (int)reflObj;
                                return (t == 2 || t == 5);
                            }
                            catch { return false; }
                        }
                        finally { try { Marshal.ReleaseComObject(reflObj); } catch { } }
                    }
                    finally { if (picture != null) try { Marshal.ReleaseComObject(picture); } catch { } }
                }
                finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }
            }
            catch { return false; }
            finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
        }

        /// <summary>4-3: スライドの画像の代替テキストを装飾化。</summary>
        public bool CheckTask_1_4_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try { slide = PowerPointCheckerCommon.GetSlideByTitle(pres, "教育理念"); } catch { }
                if (slide == null) slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
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
                            if (!PowerPointCheckerCommon.IsPictureShape(sh) && sh.Type != MsoShapeType.msoPlaceholder) continue;

                            string alt = "";
                            string title = "";
                            try { alt = sh.AlternativeText ?? ""; } catch { }
                            try { title = sh.Title ?? ""; } catch { }
                            
                            // 装飾用としてマークされているかチェック (Office 2019/365以降のフラグ)
                            bool isDecorative = false;
                            try {
                                dynamic dSh = sh;
                                // msoTrue = -1, msoFalse = 0
                                int decorativeValue = (int)dSh.Decorative;
                                if (decorativeValue == -1) isDecorative = true;
                            } catch { }

                            // 明示的に装飾用としてマークされている場合
                            if (isDecorative) return true;

                            // 互換性チェック: AltTextが空で、かつタイトルに「Decorative」というキーワードが含まれる場合など
                            if (string.IsNullOrWhiteSpace(alt) && title.IndexOf("Decorative", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                        }
                        finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                    }
                    return false;
                }
                finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
            }
            catch { return false; }
            finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
        }

        private static PptShape GetFirstPictureOnSlide(Slide slide)
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
                        bool isPic = PowerPointCheckerCommon.IsPictureShape(sh);
                        bool isPlaceholderPic = false;
                        
                        if (!isPic && sh.Type == MsoShapeType.msoPlaceholder)
                        {
                            try { if (sh.PlaceholderFormat.ContainedType == MsoShapeType.msoPicture) isPlaceholderPic = true; } catch { }
                            if (!isPlaceholderPic) {
                                string name = sh.Name.ToLower();
                                if (name.Contains("picture") || name.Contains("画像") || name.Contains("図")) isPlaceholderPic = true;
                            }
                        }

                        if (isPic || isPlaceholderPic)
                        {
                            PptShape result = sh;
                            sh = null;
                            return result;
                        }
                    }
                    finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                }
                return null;
            }
            finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
        }

        public bool CheckTask_1_4_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, "THANK YOU");
                    if (slide == null) slide = PowerPointCheckerCommon.GetSlideByNumber(pres, pres.Slides.Count); 
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
                                if (!PowerPointCheckerCommon.IsPictureShape(sh)) continue;
                                float right = (float)sh.Left + (float)sh.Width;
                                if (right > maxRight)
                                {
                                    maxRight = right;
                                    if (rightmostPicture != null) Marshal.ReleaseComObject(rightmostPicture);
                                    rightmostPicture = sh;
                                    sh = null;
                                }
                            }
                            finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                        }
                        if (rightmostPicture == null) return false;
                        try
                        {
                            dynamic pf = rightmostPicture.PictureFormat;
                            float cr = (float)pf.CropRight;
                            float cl = (float)pf.CropLeft;
                            float ct = (float)pf.CropTop;
                            float cb = (float)pf.CropBottom;
                            Marshal.ReleaseComObject(pf);
                            return cr > 0.1 || cl > 0.1 || ct > 0.1 || cb > 0.1;
                        }
                        finally { if (rightmostPicture != null) Marshal.ReleaseComObject(rightmostPicture); }
                    }
                    finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
                }
                finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }
            }
            catch { return false; }
            finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
        }

        public bool CheckTask_1_4_05()
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
                            PptShape sh = shapes[i];
                            if (!PowerPointCheckerCommon.IsPictureShape(sh)) { Marshal.ReleaseComObject(sh); continue; }
                            float left = (float)sh.Left;
                            if (left < minLeft) {
                                minLeft = left;
                                if (leftPic != null) Marshal.ReleaseComObject(leftPic);
                                leftPic = sh;
                            } else if (left > maxLeft) {
                                maxLeft = left;
                                if (rightPic != null) Marshal.ReleaseComObject(rightPic);
                                rightPic = sh;
                            } else {
                                Marshal.ReleaseComObject(sh);
                            }
                        }
                        if (leftPic == null || rightPic == null) return false;
                        float leftTop = (float)leftPic.Top;
                        float rightTop = (float)rightPic.Top;
                        bool ok = Math.Abs(rightTop - leftTop) <= PositionTolerance;
                        Marshal.ReleaseComObject(leftPic);
                        Marshal.ReleaseComObject(rightPic);
                        return ok;
                    }
                    finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
                }
                finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }
            }
            catch { return false; }
            finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
        }

        public bool CheckTask_1_4_06()
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
                        
                        var list = new List<PptShape>();
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PptShape sh = shapes[i];
                            string name = sh.Name;
                            if (name.Contains("Round Diagonal Corner Rectangle") || name.Contains("角丸斜め")) {
                                list.Add(sh);
                            } else {
                                Marshal.ReleaseComObject(sh);
                            }
                        }

                        // 右端から順に特定
                        list.Sort((a, b) => b.Left.CompareTo(a.Left));

                        if (list.Count < 3) {
                            foreach (var sh in list) Marshal.ReleaseComObject(sh);
                            return false;
                        }

                        PptShape shSmart = list[0];  // 右
                        PptShape shTablet = list[1]; // 中
                        PptShape shMonitor = list[2]; // 左

                        int zSmart = shSmart.ZOrderPosition;
                        int zTablet = shTablet.ZOrderPosition;
                        int zMonitor = shMonitor.ZOrderPosition;

                        bool ok = zSmart > zTablet && zTablet > zMonitor;

                        Marshal.ReleaseComObject(shSmart);
                        Marshal.ReleaseComObject(shTablet);
                        Marshal.ReleaseComObject(shMonitor);
                        return ok;
                    }
                    finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
                }
                finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }
            }
            catch { return false; }
            finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
        }
    }
}
