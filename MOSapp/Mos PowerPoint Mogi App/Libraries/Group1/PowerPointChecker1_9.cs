using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    public class PowerPointChecker1_9
    {
        private const int XlBarClustered = 57;
        private const string ExpectedHyperlinkAddressTask9_6 = "https://www.jica.go.jp/activities/issues/natural_env/index.html";

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
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 2);
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
                                if (sh.HasChart != MsoTriState.msoTrue) continue;
                                Chart chart = null;
                                try
                                {
                                    chart = sh.Chart;
                                    if (chart == null) continue;
                                    try
                                    {
                                        int ct = (int)chart.ChartType;
                                        return ct == XlBarClustered;
                                    }
                                    finally
                                    {
                                        if (chart != null) { try { Marshal.ReleaseComObject(chart); } catch { } }
                                    }
                                }
                                catch { continue; }
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

        public bool CheckTask_1_9_02() { return false; }
        public bool CheckTask_1_9_03() { return false; }
        public bool CheckTask_1_9_04() { return false; }
        public bool CheckTask_1_9_05() { return false; }

        public bool CheckTask_1_9_06()
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
                    PptShape shape = null;
                    try
                    {
                        shape = PowerPointCheckerCommon.FindShapeWithText(slide, "お問い合わせ");
                        if (shape == null) return false;
                        try
                        {
                            Microsoft.Office.Interop.PowerPoint.TextFrame tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)shape.TextFrame;
                            if (tf == null) return false;
                            TextRange tr = null;
                            try
                            {
                                tr = tf.TextRange;
                                if (tr == null) return false;
                                ActionSettings acts = null;
                                try
                                {
                                    acts = tr.ActionSettings;
                                    if (acts == null) return false;
                                    ActionSetting act = null;
                                    try
                                    {
                                        act = acts[PpMouseActivation.ppMouseClick];
                                        if (act == null) return false;
                                        if (act.Action != PpActionType.ppActionHyperlink) return false;
                                        Hyperlink hyp = null;
                                        try
                                        {
                                            hyp = act.Hyperlink;
                                            if (hyp != null)
                                            {
                                                string addr = (hyp.Address ?? "").Trim();
                                                if (string.Equals(addr, ExpectedHyperlinkAddressTask9_6, StringComparison.OrdinalIgnoreCase))
                                                    return true;
                                            }
                                        }
                                        finally
                                        {
                                            if (hyp != null) { try { Marshal.ReleaseComObject(hyp); } catch { } }
                                        }
                                    }
                                    finally
                                    {
                                        if (act != null) { try { Marshal.ReleaseComObject(act); } catch { } }
                                    }
                                }
                                finally
                                {
                                    if (acts != null) { try { Marshal.ReleaseComObject(acts); } catch { } }
                                }
                                try
                                {
                                    int r = 1;
                                    while (true)
                                    {
                                        TextRange run = null;
                                        try
                                        {
                                            run = tr.Runs(r, 1);
                                            if (run == null) break;
                                            ActionSettings runActs = null;
                                            try
                                            {
                                                runActs = run.ActionSettings;
                                                if (runActs == null) { r++; continue; }
                                                ActionSetting runAct = null;
                                                try
                                                {
                                                    runAct = runActs[PpMouseActivation.ppMouseClick];
                                                    if (runAct != null && runAct.Action == PpActionType.ppActionHyperlink)
                                                    {
                                                        Hyperlink runHyp = null;
                                                        try
                                                        {
                                                            runHyp = runAct.Hyperlink;
                                                            if (runHyp != null)
                                                            {
                                                                string addr = (runHyp.Address ?? "").Trim();
                                                                if (string.Equals(addr, ExpectedHyperlinkAddressTask9_6, StringComparison.OrdinalIgnoreCase))
                                                                    return true;
                                                            }
                                                        }
                                                        finally
                                                        {
                                                            if (runHyp != null) { try { Marshal.ReleaseComObject(runHyp); } catch { } }
                                                        }
                                                    }
                                                }
                                                finally
                                                {
                                                    if (runAct != null) { try { Marshal.ReleaseComObject(runAct); } catch { } }
                                                }
                                            }
                                            finally
                                            {
                                                if (runActs != null) { try { Marshal.ReleaseComObject(runActs); } catch { } }
                                            }
                                        }
                                        finally
                                        {
                                            if (run != null) { try { Marshal.ReleaseComObject(run); } catch { } }
                                        }
                                        r++;
                                    }
                                }
                                catch { }
                                return false;
                            }
                            finally
                            {
                                if (tr != null) { try { Marshal.ReleaseComObject(tr); } catch { } }
                            }
                        }
                        finally
                        {
                            if (shape != null) { try { Marshal.ReleaseComObject(shape); } catch { } }
                        }
                    }
                    catch { return false; }
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_9_07() { return false; }
    }
}
