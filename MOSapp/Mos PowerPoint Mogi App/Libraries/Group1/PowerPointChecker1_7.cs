using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using Libraries;

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

        /// <summary>7-2: 英検とはからスライド再利用。3枚目に「英検5級」を含むスライドがあるかで検証。</summary>
        public bool CheckTask_1_7_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide3 = null;
                try
                {
                    slide3 = PowerPointCheckerCommon.GetSlideByNumber(pres, 3);
                    if (slide3 == null) return false;
                    PptShape sh = null;
                    try
                    {
                        sh = PowerPointCheckerCommon.FindShapeWithText(slide3, "英検5級");
                        return sh != null;
                    }
                    finally { if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } } }
                }
                finally { if (slide3 != null) { try { Marshal.ReleaseComObject(slide3); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>7-3: アウトラインからスライド挿入。6枚目に「弊社の他の講座一覧」を含むスライドがあるかで検証。</summary>
        public bool CheckTask_1_7_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide6 = null;
                try
                {
                    slide6 = PowerPointCheckerCommon.GetSlideByNumber(pres, 6);
                    if (slide6 == null) return false;
                    PptShape sh = null;
                    try
                    {
                        sh = PowerPointCheckerCommon.FindShapeWithText(slide6, "弊社の他の講座一覧");
                        return sh != null;
                    }
                    finally { if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } } }
                }
                finally { if (slide6 != null) { try { Marshal.ReleaseComObject(slide6); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

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
                        // 「自動プレゼンテーション（Kioskモード）」に設定されているかを判定します
                        return ssSettings.ShowType == PpSlideShowType.ppShowTypeKiosk;
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
