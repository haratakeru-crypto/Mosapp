using System;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace Libraries.Group1
{
    public class PowerPointChecker1_3
    {
        public bool CheckTask_1_3_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                PptShape saShape = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    if (slide == null) return false;
                    saShape = PowerPointCheckerCommon.FindSmartArtShape(slide);
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
                            var sb = new StringBuilder();
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
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_3_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                dynamic app = null;
                try
                {
                    app = pres.Application;
                    if (app == null) return false;
                }
                catch { return false; }
                Slide slide = null;
                PptShape saShape = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    if (slide == null) return false;
                    saShape = PowerPointCheckerCommon.FindSmartArtShape(slide);
                    if (saShape == null) return false;
                    SmartArt smartArt = null;
                    try
                    {
                        smartArt = saShape.SmartArt;
                        if (smartArt == null) return false;
                        try
                        {
                            var appliedColor = smartArt.Color;
                            if (appliedColor == null) return false;
                            try
                            {
                                var accent5 = app.SmartArtColors[14];
                                if (accent5 != null && appliedColor.Equals(accent5)) return true;
                            }
                            catch { }
                            for (int idx = 1; idx <= 20; idx++)
                            {
                                try
                                {
                                    var style = app.SmartArtColors[idx];
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
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_3_03() { return false; }
        public bool CheckTask_1_3_04() { return false; }
    }
}
