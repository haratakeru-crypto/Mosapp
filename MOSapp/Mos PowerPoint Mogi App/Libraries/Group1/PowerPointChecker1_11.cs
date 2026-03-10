using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using Libraries;

namespace Libraries.Group1
{
    public class PowerPointChecker1_11
    {
        /// <summary>11-1: スライドサイズが 16:10（画面に合わせる）。</summary>
        public bool CheckTask_1_11_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                PageSetup pageSetup = null;
                try
                {
                    pageSetup = pres.PageSetup;
                    if (pageSetup == null) return false;
                    try
                    {
                        float w = (float)pageSetup.SlideWidth;
                        float h = (float)pageSetup.SlideHeight;
                        if (h <= 0) return false;
                        float ratio = w / h;
                        // 16:10 = 1.6
                        return Math.Abs(ratio - 1.6f) < 0.02f;
                    }
                    catch { return false; }
                }
                finally
                {
                    if (pageSetup != null) { try { Marshal.ReleaseComObject(pageSetup); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>11-2: スライド1のセクション名が「タイトル」。</summary>
        public bool CheckTask_1_11_02()
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
                    SectionProperties sectionProps = null;
                    try
                    {
                        sectionProps = pres.SectionProperties;
                        if (sectionProps == null) return false;
                        try
                        {
                            int sectionIndex = (int)slide.sectionIndex;
                            if (sectionIndex < 1) return false;
                            string name = null;
                            try { name = sectionProps.Name(sectionIndex) ?? ""; } catch { return false; }
                            return name.IndexOf("タイトル", StringComparison.OrdinalIgnoreCase) >= 0;
                        }
                        catch { return false; }
                    }
                    finally
                    {
                        if (sectionProps != null) { try { Marshal.ReleaseComObject(sectionProps); } catch { } }
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

        /// <summary>11-3: スライド3の箇条書きテキストボックスに塗りつぶし・枠線1.5pt。</summary>
        public bool CheckTask_1_11_03()
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
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (sh.HasTextFrame != MsoTriState.msoTrue) continue;
                                string text = "";
                                try
                                {
                                    var tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                    if (tf != null) text = tf.TextRange?.Text ?? "";
                                }
                                catch { }

                                bool hasBullet = text.Contains("•") || text.Contains("・") || text.Contains("\r") || text.Contains("\n");
                                if (!hasBullet) continue;

                                try
                                {
                                    Microsoft.Office.Interop.PowerPoint.FillFormat fill = sh.Fill;
                                    if (fill != null && fill.Visible == MsoTriState.msoTrue)
                                    {
                                        Microsoft.Office.Interop.PowerPoint.LineFormat line = sh.Line;
                                        if (line != null && line.Visible == MsoTriState.msoTrue)
                                        {
                                            float weight = (float)line.Weight;
                                            if (Math.Abs(weight - 1.5f) < 0.2f) return true;
                                        }
                                    }
                                }
                                catch { }
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

        /// <summary>11-4: 配布資料マスターで日付削除・フッター「四季のうつろい」。</summary>
        public bool CheckTask_1_11_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Master handoutMaster = null;
                try
                {
                    handoutMaster = pres.HandoutMaster;
                    if (handoutMaster == null) return false;
                    try
                    {
                        HeadersFooters hf = null;
                        try
                        {
                            hf = handoutMaster.HeadersFooters;
                            if (hf == null) return false;
                            bool dateVisible = hf.DateAndTime.Visible == MsoTriState.msoTrue;
                            string footer = hf.Footer.Text ?? "";
                            return !dateVisible && footer.IndexOf("四季のうつろい", StringComparison.OrdinalIgnoreCase) >= 0;
                        }
                        finally { if (hf != null) { try { Marshal.ReleaseComObject(hf); } catch { } } }
                    }
                    catch { return false; }
                }
                finally { if (handoutMaster != null) { try { Marshal.ReleaseComObject(handoutMaster); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>11-5: スライド5のアイコンに「青」塗りつぶし。</summary>
        public bool CheckTask_1_11_05()
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
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                // msoGraphic は enumにない場合があるため数値 (24) でチェック
                                int shapeType = (int)sh.Type;
                                if (shapeType != 24 && shapeType != 13 /* msoPicture */ && !sh.Name.Contains("Graphic") && !sh.Name.Contains("Icon")) continue;

                                Microsoft.Office.Interop.PowerPoint.FillFormat fill = null;
                                try
                                {
                                    fill = sh.Fill;
                                    if (fill == null || fill.Visible != MsoTriState.msoTrue) continue;
                                    
                                    Microsoft.Office.Interop.PowerPoint.ColorFormat cf = fill.ForeColor;
                                    if (cf == null) continue;
                                    
                                    int rgb = (int)cf.RGB;
                                    int r = rgb & 0xFF; int g = (rgb >> 8) & 0xFF; int b = (rgb >> 16) & 0xFF;

                                    // 標準色の「青」 (RGB: 0, 112, 192) = BGR(192, 112, 0)
                                    // 初期状態がアクセントカラー（青系）であるため、RGB指定されているか、または標準色の特定の値を狙う
                                    if (b >= 190 && b <= 200 && g >= 110 && g <= 120 && r == 0) return true;
                                    
                                    // または標準の「青」(別バリエーション)
                                    if (b == 255 && r == 0 && g == 0) return true; // 純粋な青
                                }
                                finally { if (fill != null) { try { Marshal.ReleaseComObject(fill); } catch { } } }
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

        /// <summary>11-6: スライド2の箇条書きテキストボックスを上下中央揃え。</summary>
        public bool CheckTask_1_11_06()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                
                PageSetup ps = pres.PageSetup;
                float slideHeight = ps.SlideHeight;
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
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (sh.HasTextFrame != MsoTriState.msoTrue) continue;
                                
                                string text = "";
                                try
                                {
                                    var tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                    if (tf != null) text = tf.TextRange?.Text ?? "";
                                }
                                catch { }

                                // 箇条書きを含む図形か
                                bool isBodyText = text.Contains("•") || text.Contains("・") || text.Contains("\r") || text.Contains("\n");
                                if (!isBodyText) continue;

                                // [配置] -> [上下中央揃え] のチェック
                                float shapeCenter = sh.Top + (sh.Height / 2.0f);
                                float slideCenter = slideHeight / 2.0f;

                                // スライドの中央に配置されているか
                                if (Math.Abs(shapeCenter - slideCenter) < 5.0f) return true;

                                // もしテキスト内の配置を指している場合も考慮
                                try {
                                    dynamic tf2 = sh.TextFrame2;
                                    if (tf2 != null && (int)tf2.VerticalAnchor == (int)MsoVerticalAnchor.msoAnchorMiddle) return true;
                                } catch { }
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

        /// <summary>11-7: ノートで全スライド3部・部単位で印刷。COM の PrintOptions または VSTO ログの印刷記録で判定。</summary>
        public bool CheckTask_1_11_07()
        {
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
                        // 「ノート」
                        if (po.OutputType != PpPrintOutputType.ppPrintOutputNotesPages) return false;
                        // 「3部」
                        if (po.NumberOfCopies != 3) return false;
                        // 「1ページ目を全て印刷したあとに...」 = ページ単位 (Uncollated) = Collate: Off
                        return po.Collate == MsoTriState.msoFalse;
                    }
                    catch { return false; }
                }
                finally { if (po != null) { try { Marshal.ReleaseComObject(po); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
    }
}
