using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題セット1 プロジェクト1（干潟観察会）。
    /// 最終状態のみで判定する（VSTO 証跡は不要）。
    /// </summary>
    public class PowerPointChecker1_1_PV1
    {
        private const string SlideTitleColumns = "三つの観察は、干潟の奥から順に見ます";
        private const string SlideTitleAnimation = "残り13組は、5月8日の夕方までです";
        private const string CalloutText = "初めての方へ";
        private const string SlideTitleSmiley = "当日は、受付、観察、カードの順です";
        private const string CommentText = "干潟の危険は詳しく";
        private const string SlideTitleBackground = "観察は、触る前に止まる";

        private const int MsoAnimEffectBox = 4;
        private const int MsoAnimateLevelNone = 0;

        public bool CheckTask_1_1_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleColumns);
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
                                if (sh.HasTextFrame != MsoTriState.msoTrue) continue;
                                try
                                {
                                    var tf2 = sh.TextFrame2;
                                    if (tf2 == null) continue;
                                    var col = tf2.Column;
                                    if (col == null) continue;
                                    if (col.Number == 2) return true;
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
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        public bool CheckTask_1_1_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleAnimation);
                    if (slide == null) return false;

                    Sequence seq = null;
                    try
                    {
                        seq = slide.TimeLine?.MainSequence;
                        if (seq == null) return false;
                        int count = seq.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            Effect eff = null;
                            try
                            {
                                eff = seq[i];
                                if (eff == null) continue;
                                if (!IsBoxEntranceEffect(eff)) continue;
                                if (IsAsOneObjectTextEffect(eff))
                                    return true;
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
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        public bool CheckTask_1_1_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                PptShape shape = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                    if (slide == null) return false;
                    shape = PowerPointCheckerCommon.FindShapeWithText(slide, CalloutText);
                    return shape != null;
                }
                finally
                {
                    if (shape != null) { try { Marshal.ReleaseComObject(shape); } catch { } }
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        public bool CheckTask_1_1_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleSmiley);
                    if (slide == null) return false;

                    int smileys = 0;
                    int bolts = 0;
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
                                    var t = sh.AutoShapeType;
                                    if (t == MsoAutoShapeType.msoShapeSmileyFace)
                                        smileys++;
                                    else if (t == MsoAutoShapeType.msoShapeLightningBolt)
                                        bolts++;
                                }
                                catch { }
                            }
                            finally
                            {
                                if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                            }
                        }
                    }
                    finally
                    {
                        if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
                    }

                    return smileys >= 1 && bolts == 0;
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

        public bool CheckTask_1_1_05()
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
                                if (text.IndexOf(CommentText, StringComparison.OrdinalIgnoreCase) >= 0)
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
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
            }
        }

        public bool CheckTask_1_1_06()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleBackground);
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
                            finally
                            {
                                if (cf != null) { try { Marshal.ReleaseComObject(cf); } catch { } }
                            }
                        }
                        finally
                        {
                            if (fill != null) { try { Marshal.ReleaseComObject(fill); } catch { } }
                        }
                    }
                    finally
                    {
                        if (bg != null) { try { Marshal.ReleaseComObject(bg); } catch { } }
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

        private static bool IsBoxEntranceEffect(Effect eff)
        {
            if (eff == null) return false;
            try
            {
                if ((int)eff.EffectType == MsoAnimEffectBox)
                    return true;
            }
            catch { }
            try
            {
                string dn = eff.DisplayName ?? "";
                if (dn.IndexOf("ボックス", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                if (dn.IndexOf("Box", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 「1つのオブジェクトとして」: 段落単位ビルドではなく、図形全体が1効果として動く。
        /// BuildByLevel=None、または TextUnitEffect=Mixed かつ ByParagraph でないこと。
        /// </summary>
        private static bool IsAsOneObjectTextEffect(Effect eff)
        {
            if (eff == null) return false;
            EffectInformation info = null;
            try { info = eff.EffectInformation; } catch { }
            if (info != null)
            {
                try
                {
                    int buildByLevel = (int)info.BuildByLevelEffect;
                    if (buildByLevel == MsoAnimateLevelNone)
                        return true;
                }
                catch { }
                try
                {
                    if (info.TextUnitEffect == MsoAnimTextUnitEffect.msoAnimTextUnitEffectByParagraph)
                        return false;
                    if (info.TextUnitEffect == MsoAnimTextUnitEffect.msoAnimTextUnitEffectMixed)
                        return true;
                }
                catch { }
            }
            try
            {
                // 段落番号付きビルド効果ではないこと
                if (eff.Paragraph == 0)
                    return true;
            }
            catch { }
            return false;
        }
    }
}
