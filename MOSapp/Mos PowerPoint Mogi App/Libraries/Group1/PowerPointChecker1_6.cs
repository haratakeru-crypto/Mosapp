using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace Libraries.Group1
{
    public class PowerPointChecker1_6
    {
        /// <summary>6-1: スライド1にコメント「MOSの説明は詳しく」が存在するか。</summary>
        public bool CheckTask_1_6_01()
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
                    Comments comments = null;
                    try
                    {
                        comments = slide.Comments;
                        if (comments == null) return false;
                        int count = comments.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            Comment cmt = null;
                            try
                            {
                                cmt = comments[i];
                                if (cmt == null) continue;
                                string text = null;
                                try { text = cmt.Text ?? ""; } catch { }
                                if (text.IndexOf("MOSの説明は詳しく", StringComparison.OrdinalIgnoreCase) >= 0)
                                    return true;
                            }
                            finally
                            {
                                if (cmt != null) { try { Marshal.ReleaseComObject(cmt); } catch { } }
                            }
                        }
                        return false;
                    }
                    finally
                    {
                        if (comments != null) { try { Marshal.ReleaseComObject(comments); } catch { } }
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

        /// <summary>6-2: スライド3の背景を「薄い灰色 背景2」に。</summary>
        public bool CheckTask_1_6_02()
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
                    Microsoft.Office.Interop.PowerPoint.ShapeRange bg = null;
                    try
                    {
                        bg = slide.Background;
                        if (bg == null) return false;
                        Microsoft.Office.Interop.PowerPoint.FillFormat fill = null;
                        try
                        {
                            fill = bg.Fill;
                            if (fill == null) return false;
                            if (fill.Visible != MsoTriState.msoTrue) return false;
                            Microsoft.Office.Interop.PowerPoint.ColorFormat cf = null;
                            try
                            {
                                cf = fill.ForeColor;
                                if (cf == null) return false;
                                return cf.ObjectThemeColor == MsoThemeColorIndex.msoThemeColorBackground2;
                            }
                            finally { if (cf != null) { try { Marshal.ReleaseComObject(cf); } catch { } } }
                        }
                        finally { if (fill != null) { try { Marshal.ReleaseComObject(fill); } catch { } } }
                    }
                    finally { if (bg != null) { try { Marshal.ReleaseComObject(bg); } catch { } } }
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>6-3: スライド「MOSについて」に3Dモデル幅9.5。</summary>
        public bool CheckTask_1_6_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, "MOSについて");
                    if (slide == null) return false;
                    PptShape modelShape = null;
                    try
                    {
                        modelShape = PowerPointCheckerCommon.Find3DModelShape(slide);
                        if (modelShape == null) return false;
                        try
                        {
                            float w = (float)modelShape.Width;
                            float cm = w / 28.34646f;
                            return Math.Abs(cm - 9.5f) < 0.1f;
                        }
                        catch { return false; }
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

        /// <summary>6-4: スライド5の3D「Shikaku」を下背面・高さ7.2に。</summary>
        public bool CheckTask_1_6_04()
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
                    
                    PptShape modelShape = null;
                    try
                    {
                        modelShape = PowerPointCheckerCommon.Find3DModelShape(slide);
                        if (modelShape == null) return false;
                        
                        try
                        {
                            bool viewMatch = false;
                            try
                            {
                                dynamic dynShape = modelShape;
                                dynamic m3d = dynShape.Model3D;
                                if (m3d != null)
                                {
                                    // 下背面の回転データ (RotX=20, RotY=180, RotZ=0)
                                    float rotX = (float)m3d.RotationX;
                                    float rotY = (float)m3d.RotationY;
                                    float rotZ = (float)m3d.RotationZ;

                                    if (Math.Abs(rotX - 20f) < 1f && 
                                        Math.Abs(rotY - 180f) < 1f && 
                                        Math.Abs(rotZ - 0f) < 1f)
                                    {
                                        viewMatch = true;
                                    }
                                }
                            }
                            catch { }

                            float h = (float)modelShape.Height;
                            float cm = h / 28.34646f;
                            bool heightMatch = Math.Abs(cm - 7.2f) < 0.1f;
                            
                            return viewMatch && heightMatch;
                        }
                        catch { return false; }
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
    }
}
