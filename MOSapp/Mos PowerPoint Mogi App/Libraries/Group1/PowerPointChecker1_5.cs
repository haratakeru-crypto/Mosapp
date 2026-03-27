using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using Libraries;

namespace Libraries.Group1
{
    public class PowerPointChecker1_5
    {
        /// <summary>5-1: 配布資料3スライド・部単位4部印刷設定。COM の PrintOptions または VSTO ログの印刷記録で判定。</summary>
        public bool CheckTask_1_5_01()
        {
            if (PPLogReader.HasTask5_1PrintExecuted())
                return true;
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                PrintOptions po = null;
                try
                {
                    po = pres.PrintOptions;
                    if (po == null) return false;
                    try
                    {
                        if (po.OutputType != PpPrintOutputType.ppPrintOutputThreeSlideHandouts) return false;
                        if (po.NumberOfCopies != 4) return false;
                        if (po.Collate != MsoTriState.msoTrue) return false;
                        return true;
                    }
                    catch { return false; }
                }
                finally { if (po != null) { try { Marshal.ReleaseComObject(po); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>5-2: スライド5の「いつからでも」の文字塗りつぶしを「茶、テキスト2」に。</summary>
        public bool CheckTask_1_5_02()
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
                    PptShape sh = null;
                    try
                    {
                        sh = PowerPointCheckerCommon.FindShapeWithText(slide, "いつからでも");
                        if (sh == null) return false;
                        Microsoft.Office.Interop.PowerPoint.TextFrame tf = null;
                        try
                        {
                            if (sh.HasTextFrame != MsoTriState.msoTrue) return false;
                            tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                            if (tf == null) return false;
                            TextRange tr = null;
                            try
                            {
                                tr = tf.TextRange;
                                if (tr == null) return false;
                                Font font = null;
                                try
                                {
                                    font = tr.Font;
                                    if (font == null) return false;
                                    Microsoft.Office.Interop.PowerPoint.ColorFormat cf = null;
                                    try
                                    {
                                        cf = font.Color;
                                        if (cf == null) return false;
                                        return cf.ObjectThemeColor == MsoThemeColorIndex.msoThemeColorText2;
                                    }
                                    finally { if (cf != null) { try { Marshal.ReleaseComObject(cf); } catch { } } }
                                }
                                finally { if (font != null) { try { Marshal.ReleaseComObject(font); } catch { } } }
                            }
                            finally { if (tr != null) { try { Marshal.ReleaseComObject(tr); } catch { } } }
                        }
                        finally { if (tf != null) { try { Marshal.ReleaseComObject(tf); } catch { } } }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
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

        /// <summary>5-3: スライド4の稲妻図形をスマイルに変更済みか（AutoShapeType = Smiley）。</summary>
        public bool CheckTask_1_5_03()
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
                                try
                                {
                                    if (sh.AutoShapeType == MsoAutoShapeType.msoShapeSmileyFace)
                                        return true;
                                }
                                catch { }
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
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>5-4: スライド4の小さい雲の幅が他の雲と同じ。</summary>
        public bool CheckTask_1_5_04()
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
                    PptShapes shapes = null;
                    try
                    {
                        shapes = slide.Shapes;
                        if (shapes == null) return false;
                        
                        int count = shapes.Count;
                        System.Collections.Generic.List<float> cloudWidths = new System.Collections.Generic.List<float>();
                        for (int i = 1; i <= count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                bool isCloudCandidate = false;
                                // 画像であるか、またはオートシェイプの「雲」であるかを確認
                                if (PowerPointCheckerCommon.IsPictureShape(sh)) 
                                {
                                    isCloudCandidate = true;
                                }
                                else 
                                {
                                    try 
                                    { 
                                        if (sh.AutoShapeType == Microsoft.Office.Core.MsoAutoShapeType.msoShapeCloud) 
                                            isCloudCandidate = true; 
                                    } catch { }
                                }
                                if (isCloudCandidate)
                                {
                                    cloudWidths.Add((float)sh.Width);
                                }
                            }
                            finally
                            {
                                if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                            }
                        }
                        // 画像または雲の図形が2つ以上ない場合は不正解
                        // 画像または雲の図形が3つ以上ない場合は不正解（対象となる雲が3つあるため）
                        if (cloudWidths.Count < 3) return false;
                        
                        // 幅を降順（大きい順）に並び替え
                        // 初期の「大きい雲2つ＋小さい雲1つ」の状態では、
                        // 1番目(大きい雲)と3番目(小さい雲)で幅が異なるため不正解になる。
                        // 正しく修正すると、雲3つの幅が揃うため、1番目と3番目の幅が一致して正解となる。
                        cloudWidths.Sort();
                        cloudWidths.Reverse();
                        
                        // 上位3つの内、最大と最小の差分が0.5未満か確認
                        return Math.Abs(cloudWidths[0] - cloudWidths[2]) < 0.5f;
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

        /// <summary>5-5: スライド3の3つの図形がグループ化されているか。</summary>
        public bool CheckTask_1_5_05()
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
                        int count = shapes.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (sh.Type != MsoShapeType.msoGroup) continue;
                                Microsoft.Office.Interop.PowerPoint.GroupShapes group = null;
                                try
                                {
                                    group = sh.GroupItems;
                                    if (group != null && group.Count >= 3)
                                        return true;
                                }
                                finally
                                {
                                    if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
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
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
    }
}
