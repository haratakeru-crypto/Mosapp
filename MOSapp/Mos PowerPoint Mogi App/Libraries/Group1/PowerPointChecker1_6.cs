using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Libraries;

namespace Libraries.Group1
{
    /// <summary>プロジェクト6（P6-1〜P6-7）。Phase B でタスク単位に Legacy 移植。</summary>
    public class PowerPointChecker1_6
    {
        /// <summary>P6-1: ドキュメント検査 — 全スライドコメント0件、プロパティ・個人情報が空（旧10-1）。</summary>
        public bool CheckTask_1_6_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                Slides slides = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null) return false;
                    for (int i = 1; i <= slides.Count; i++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            Comments comments = null;
                            try
                            {
                                comments = slide.Comments;
                                if (comments != null && comments.Count > 0) return false;
                            }
                            finally { if (comments != null) { try { Marshal.ReleaseComObject(comments); } catch { } } }
                        }
                        finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
                    }
                }
                finally { if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } } }

                try
                {
                    dynamic props = pres.BuiltInDocumentProperties;
                    if (props == null) return true;
                    string[] personalPropNames = { "Author", "Manager", "Company", "Last Author", "Title", "Subject", "Keywords", "Comments" };
                    foreach (string name in personalPropNames)
                    {
                        try
                        {
                            object val = props[name].Value;
                            string s = (val == null) ? "" : (val.ToString() ?? "").Trim();
                            if (!string.IsNullOrEmpty(s)) return false;
                        }
                        catch { /* プロパティが存在しない場合は無視 */ }
                    }
                    return true;
                }
                catch { return false; }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P6-2: 常に読み取り専用 — pptx 内 readOnlyRecommended または ReadOnly（旧8-5）。</summary>
        public bool CheckTask_1_6_02()
        {
            Presentation pres = null;
            string tempPath = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                tempPath = Path.Combine(Path.GetTempPath(), "mos_6_2_check_" + Guid.NewGuid().ToString("N") + ".pptx");
                try
                {
                    pres.SaveCopyAs(tempPath);
                }
                catch { return false; }

                try
                {
                    using (var package = Package.Open(tempPath, FileMode.Open, FileAccess.Read))
                    {
                        foreach (var part in package.GetParts())
                        {
                            if (part.Uri.OriginalString.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                            {
                                string xml;
                                using (var reader = new StreamReader(part.GetStream()))
                                    xml = reader.ReadToEnd();

                                if (Regex.IsMatch(xml, @"readOnlyRecommended[^>]*val\s*=\s*""(?:1|true)""", RegexOptions.IgnoreCase))
                                    return true;
                            }
                        }
                    }
                }
                catch { }

                try
                {
                    if (pres.ReadOnly == MsoTriState.msoTrue) return true;
                }
                catch { }

                return false;
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
                if (tempPath != null && File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
        }

        /// <summary>P6-3: スライドショー自動プレゼンテーション（Kiosk）。VSTO 証跡または COM（旧7-4）。</summary>
        public bool CheckTask_1_6_03()
        {
            if (PPLogReader.HasTask6_3KioskExecuted())
                return true;
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

        /// <summary>P6-4: 目的別スライドショー「書式のポイント」にスライド4・5・6がこの順で含まれる（旧10-2、名称差分）。</summary>
        public bool CheckTask_1_6_04()
        {
            const string expectedShowName = "書式のポイント";
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                int[] expectedSlideIds = new int[3];
                for (int k = 0; k < 3; k++)
                {
                    Slide slide = null;
                    try
                    {
                        slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 4 + k);
                        if (slide == null) return false;
                        expectedSlideIds[k] = slide.SlideID;
                    }
                    finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
                }

                SlideShowSettings ssSettings = null;
                try
                {
                    ssSettings = pres.SlideShowSettings;
                    if (ssSettings == null) return false;
                    NamedSlideShows namedShows = null;
                    try
                    {
                        namedShows = ssSettings.NamedSlideShows;
                        if (namedShows == null) return false;
                        int count = namedShows.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            NamedSlideShow ns = null;
                            try
                            {
                                ns = namedShows[i];
                                if (ns == null) continue;
                                string name = null;
                                try { name = ns.Name ?? ""; } catch { continue; }
                                if (!string.Equals((name ?? "").Trim(), expectedShowName, StringComparison.Ordinal))
                                    continue;
                                int[] showIds = TryGetNamedSlideShowSlideIds(ns);
                                if (showIds == null || showIds.Length != 3) continue;
                                if (showIds[0] == expectedSlideIds[0]
                                    && showIds[1] == expectedSlideIds[1]
                                    && showIds[2] == expectedSlideIds[2])
                                    return true;
                            }
                            finally
                            {
                                if (ns != null) { try { Marshal.ReleaseComObject(ns); } catch { } }
                            }
                        }
                        return false;
                    }
                    finally
                    {
                        if (namedShows != null) { try { Marshal.ReleaseComObject(namedShows); } catch { } }
                    }
                }
                finally
                {
                    if (ssSettings != null) { try { Marshal.ReleaseComObject(ssSettings); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        private static int[] NormalizeSlideIdsFromComArray(Array arr)
        {
            var list = new List<int>();
            for (int i = 0; i < arr.Length; i++)
            {
                object o = arr.GetValue(i);
                if (o == null || Equals(o, Missing.Value))
                    continue;
                try
                {
                    int id = Convert.ToInt32(o);
                    if (id == 0)
                        continue;
                    list.Add(id);
                }
                catch { }
            }
            return list.ToArray();
        }

        private static int[] TryGetNamedSlideShowSlideIds(NamedSlideShow namedShow)
        {
            if (namedShow == null) return null;
            object raw = null;
            try
            {
                try { raw = namedShow.SlideIDs; }
                catch { return null; }
                if (raw == null) return null;
                if (raw is Array arr)
                {
                    if (arr.Length == 0) return new int[0];
                    return NormalizeSlideIdsFromComArray(arr);
                }
                return new[] { Convert.ToInt32(raw) };
            }
            catch
            {
                return null;
            }
            finally
            {
                if (raw != null && Marshal.IsComObject(raw))
                {
                    try { Marshal.ReleaseComObject(raw); } catch { }
                }
            }
        }

        /// <summary>P6-5: アウトライン・部単位6部印刷。COM の PrintOptions または VSTO ログの印刷記録で判定（旧5-1、条件差分）。</summary>
        public bool CheckTask_1_6_05()
        {
            if (PPLogReader.HasTask6_5PrintExecuted())
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
                        if (po.OutputType != PpPrintOutputType.ppPrintOutputOutline) return false;
                        if (po.NumberOfCopies != 6) return false;
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

        /// <summary>P6-6: ノート3部・ページ単位印刷（Collate OFF）。COM または VSTO 証跡で判定（旧11-7、Collate 条件は異なる）。</summary>
        public bool CheckTask_1_6_06()
        {
            if (PPLogReader.HasTask6_6PrintExecuted())
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
                        if (po.OutputType != PpPrintOutputType.ppPrintOutputNotesPages) return false;
                        if (po.NumberOfCopies != 3) return false;
                        if (po.Collate != MsoTriState.msoFalse) return false;
                        return true;
                    }
                    catch { return false; }
                }
                finally { if (po != null) { try { Marshal.ReleaseComObject(po); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>P6-7: グレースケール配布資料3スライド/頁・4部。COM（PrintColorType=ppPrintBlackAndWhite）または VSTO 証跡で判定。</summary>
        public bool CheckTask_1_6_07()
        {
            if (PPLogReader.HasTask6_7PrintExecuted())
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
                        if (po.PrintColorType != PpPrintColorType.ppPrintBlackAndWhite) return false;
                        return true;
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
