using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace Libraries.Group1
{
    public class PowerPointChecker1_7
    {
        /// <summary>7-1: スライド2～5がセクション「資格情報」に属するか。</summary>
        public bool CheckTask_1_7_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                SectionProperties sp = null;
                try
                {
                    sp = pres.SectionProperties;
                    if (sp == null) return false;
                    int sectionCount = sp.Count;
                    int sectionIndexForQual = -1;
                    for (int s = 1; s <= sectionCount; s++)
                    {
                        try
                        {
                            string name = sp.Name(s) ?? "";
                            if (name.IndexOf("資格情報", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                sectionIndexForQual = s;
                                break;
                            }
                        }
                        catch { }
                    }
                    if (sectionIndexForQual < 1) return false;
                    for (int slideNum = 2; slideNum <= 5; slideNum++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = PowerPointCheckerCommon.GetSlideByNumber(pres, slideNum);
                            if (slide == null) return false;
                            try
                            {
                                int sid = (int)slide.sectionIndex;
                                if (sid != sectionIndexForQual) return false;
                            }
                            catch { return false; }
                        }
                        finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
                    }
                    return true;
                }
                finally
                {
                    if (sp != null) { try { Marshal.ReleaseComObject(sp); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        public bool CheckTask_1_7_02() { return false; }
        public bool CheckTask_1_7_03() { return false; }

        /// <summary>7-4: スライドショーを自動プレゼンテーションに設定。</summary>
        public bool CheckTask_1_7_04()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                SlideShowSettings ssSettings = null;
                try
                {
                    ssSettings = pres.SlideShowSettings;
                    if (ssSettings == null) return false;
                    try
                    {
                        return ssSettings.AdvanceMode == PpSlideShowAdvanceMode.ppSlideShowUseSlideTimings;
                    }
                    catch { return false; }
                }
                finally
                {
                    if (ssSettings != null) { try { Marshal.ReleaseComObject(ssSettings); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }
    }
}
