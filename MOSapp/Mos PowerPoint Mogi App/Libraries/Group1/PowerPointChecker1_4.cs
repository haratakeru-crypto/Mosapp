using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_4
    {
        private const double PositionTolerance = 2.0;

        public bool CheckTask_1_4_01() { return false; }
        public bool CheckTask_1_4_02() { return false; }
        public bool CheckTask_1_4_03() { return false; }

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
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
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
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (!PowerPointCheckerCommon.IsPictureShape(sh)) continue;
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
                                if (!PowerPointCheckerCommon.IsPictureShape(sh)) continue;
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
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
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
