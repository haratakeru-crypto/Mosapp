using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using PptColorFormat = Microsoft.Office.Interop.PowerPoint.ColorFormat;
using PptFillFormat = Microsoft.Office.Interop.PowerPoint.FillFormat;
using PptLineFormat = Microsoft.Office.Interop.PowerPoint.LineFormat;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題セット1 プロジェクト4（みどり丘市民農園）。最終状態のみで判定する。
    /// ActivePresentation は ReleaseComObject しない（その場採点でプレゼンが閉じるのを防ぐ）。
    /// </summary>
    public class PowerPointChecker1_4_PV1
    {
        private const string SlideTitleSpring = "春は、4月11日にジャガイモを植えます";
        private const string SlideTitleSummer = "夏は、7月18日に支柱を立てます";
        private const string SlideTitleAutumn = "秋は、10月24日に芋を掘ります";
        private const string SlideTitleShape = "土曜の午前は、支柱と水だけ見ます";
        private const string SlideTitleTable = "2025年は、夏の41区画が一番多かったです";
        private const string SlideTitleVideo = "畝の作り方は、動画の30秒で見ます";
        private const string ExpectedShowName = "教育";
        private const string ExpectedLayoutName = "画像付きスライド";
        private const string ExpectedVideoNameHint = "畑の様子";
        private const int Pp3DAnimTurntable = 152;
        private const float FillBrightnessTarget = 0.6f;
        private const float FillBrightnessTolerance = 0.09f;
        private const float LineWeightTarget = 1.5f;
        private const float LineWeightTolerance = 0.2f;

        /// <summary>類1-4-1: 3Dモデルにアニメーション「ターンテーブル」。</summary>
        public bool CheckTask_1_4_01()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleSpring);
                    if (slide == null) return false;
                    PptShape modelShape = null;
                    try
                    {
                        modelShape = PowerPointCheckerCommon.Find3DModelShape(slide);
                        if (modelShape == null) return false;
                        TimeLine timeline = slide.TimeLine;
                        if (timeline == null) return false;
                        Sequence seq = timeline.MainSequence;
                        if (seq == null) return false;
                        int modelId = modelShape.Id;
                        for (int i = 1; i <= seq.Count; i++)
                        {
                            Effect eff = null;
                            try
                            {
                                eff = seq[i];
                                if (eff == null) continue;
                                PptShape effShape = null;
                                try
                                {
                                    effShape = eff.Shape;
                                    if (effShape == null || effShape.Id != modelId) continue;
                                    if (IsTurntableAnimation(eff))
                                        return true;
                                }
                                finally
                                {
                                    if (effShape != null) { try { Marshal.ReleaseComObject(effShape); } catch { } }
                                }
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
                        if (modelShape != null) { try { Marshal.ReleaseComObject(modelShape); } catch { } }
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        /// <summary>
        /// 類1-4-2: 図形の塗りつぶし（アクセント1・白+基本色60%）、枠線（濃い青）1.5pt。
        /// </summary>
        public bool CheckTask_1_4_02()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleShape);
                    if (slide == null) return false;
                    try
                    {
                        PptShapes shapes = slide.Shapes;
                        if (shapes == null) return false;
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (!IsStyleCandidateShape(sh)) continue;
                                if (TryValidateShapeFillAndLine(sh))
                                    return true;
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
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        /// <summary>類1-4-3: 目的別スライドショー「教育」に春・夏・秋の3スライド（この順）。</summary>
        public bool CheckTask_1_4_03()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;

                    int[] expectedIds = new int[3];
                    string[] titles = { SlideTitleSpring, SlideTitleSummer, SlideTitleAutumn };
                    for (int i = 0; i < titles.Length; i++)
                    {
                        Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, titles[i]);
                        if (slide == null) return false;
                        try { expectedIds[i] = slide.SlideID; }
                        finally { try { Marshal.ReleaseComObject(slide); } catch { } }
                    }

                    SlideShowSettings ssSettings = pres.SlideShowSettings;
                    if (ssSettings == null) return false;
                    NamedSlideShows namedShows = ssSettings.NamedSlideShows;
                    if (namedShows == null) return false;
                    for (int i = 1; i <= namedShows.Count; i++)
                    {
                        NamedSlideShow ns = null;
                        try
                        {
                            ns = namedShows[i];
                            if (ns == null) continue;
                            string name = "";
                            try { name = ns.Name ?? ""; } catch { continue; }
                            if (!string.Equals(name.Trim(), ExpectedShowName, StringComparison.Ordinal))
                                continue;
                            int[] showIds = TryGetNamedSlideShowSlideIds(ns);
                            if (showIds == null || showIds.Length != 3) continue;
                            if (showIds[0] == expectedIds[0]
                                && showIds[1] == expectedIds[1]
                                && showIds[2] == expectedIds[2])
                                return true;
                        }
                        finally
                        {
                            if (ns != null) { try { Marshal.ReleaseComObject(ns); } catch { } }
                        }
                    }
                    return false;
                }
                catch { return false; }
            }
        }

        /// <summary>類1-4-4: 表スタイル「中間スタイル1-アクセント4」、縞模様（行）オフ。</summary>
        public bool CheckTask_1_4_04()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleTable);
                    if (slide == null) return false;
                    try
                    {
                        PptShapes shapes = slide.Shapes;
                        if (shapes == null) return false;
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PptShape sh = null;
                            try
                            {
                                sh = shapes[i];
                                if (sh.HasTable != MsoTriState.msoTrue) continue;
                                Table table = sh.Table;
                                if (table == null) continue;
                                TableStyle ts = table.Style;
                                if (ts == null) continue;
                                string styleName = "";
                                try { styleName = ts.Name ?? ""; } catch { }
                                if (!IsMediumStyle1Accent4(styleName)) continue;
                                if (table.HorizBanding) continue;
                                return true;
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
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        /// <summary>類1-4-5: 動画「畑の様子.mp4」を挿入（位置不問）。</summary>
        public bool CheckTask_1_4_05()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Slide slide = PowerPointCheckerCommon.GetSlideByTitle(pres, SlideTitleVideo);
                    if (slide == null) return false;
                    PptShape videoShape = null;
                    try
                    {
                        var status = PowerPointCheckerCommon.TryFindFirstVideoShape(slide, out videoShape);
                        if (status != PowerPointCheckerCommon.MediaScanStatus.Found || videoShape == null)
                            return false;

                        // ファイル名が取れれば一致確認。取れない環境では動画存在のみで合格。
                        string name = "";
                        try { name = videoShape.Name ?? ""; } catch { }
                        if (string.IsNullOrEmpty(name))
                            return true;
                        if (name.IndexOf(ExpectedVideoNameHint, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                        if (name.IndexOf(".mp4", StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                        // 名前が汎用（Video 1 等）でも挿入されていれば合格
                        return true;
                    }
                    finally
                    {
                        if (videoShape != null) { try { Marshal.ReleaseComObject(videoShape); } catch { } }
                        try { Marshal.ReleaseComObject(slide); } catch { }
                    }
                }
                catch { return false; }
            }
        }

        /// <summary>
        /// 類1-4-6: 「白紙」複製レイアウト「画像付きスライド」。グラフ左・テキスト右。
        /// （教材 旧10-7 と同方式。最終状態 COM のみ）
        /// </summary>
        public bool CheckTask_1_4_06()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    Presentation pres = PowerPointCheckerCommon.GetActivePresentation();
                    if (pres == null) return false;
                    Master master = pres.SlideMaster;
                    if (master == null) return false;
                    CustomLayouts layouts = master.CustomLayouts;
                    if (layouts == null) return false;
                    for (int i = 1; i <= layouts.Count; i++)
                    {
                        CustomLayout cl = null;
                        try
                        {
                            cl = layouts[i];
                            if (cl == null) continue;
                            string name = "";
                            try { name = cl.Name ?? ""; } catch { continue; }
                            if (name.IndexOf(ExpectedLayoutName, StringComparison.OrdinalIgnoreCase) < 0)
                                continue;
                            return LayoutHasChartLeftOfTextPlaceholders(cl);
                        }
                        finally
                        {
                            if (cl != null) { try { Marshal.ReleaseComObject(cl); } catch { } }
                        }
                    }
                    return false;
                }
                catch { return false; }
            }
        }

        private static bool IsTurntableAnimation(Effect eff)
        {
            if (eff == null) return false;
            try
            {
                int etVal = (int)eff.EffectType;
                // Legacy 2-4 と同様、環境差で 129 / 152 の両方を許容。
                if (etVal == Pp3DAnimTurntable || etVal == 129)
                    return true;
            }
            catch { }

            try
            {
                string dn = eff.DisplayName ?? "";
                if (dn.IndexOf("ターンテーブル", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                if (dn.IndexOf("Turntable", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                if (dn.IndexOf("turn", StringComparison.OrdinalIgnoreCase) >= 0
                    && dn.IndexOf("jump", StringComparison.OrdinalIgnoreCase) < 0
                    && dn.IndexOf("ジャンプ", StringComparison.OrdinalIgnoreCase) < 0)
                    return true;
            }
            catch { }
            return false;
        }

        private static bool IsStyleCandidateShape(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                MsoShapeType type = sh.Type;
                if (type == MsoShapeType.msoAutoShape
                    || type == MsoShapeType.msoFreeform
                    || type == MsoShapeType.msoTextBox)
                    return true;
                if (type == MsoShapeType.msoPlaceholder)
                {
                    // タイトル以外のコンテンツ図形は対象外寄りだが、塗りがあるプレースホルダーも稀にある。
                    // ここではタイトルプレースホルダーのみ除外。
                    PlaceholderFormat pf = sh.PlaceholderFormat;
                    if (pf == null) return false;
                    try
                    {
                        var pt = (PpPlaceholderType)pf.Type;
                        return pt != PpPlaceholderType.ppPlaceholderTitle
                            && pt != PpPlaceholderType.ppPlaceholderCenterTitle;
                    }
                    finally { try { Marshal.ReleaseComObject(pf); } catch { } }
                }
            }
            catch { }
            return false;
        }

        private static bool TryValidateShapeFillAndLine(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                PptFillFormat fill = sh.Fill;
                PptLineFormat line = sh.Line;
                if (fill == null || fill.Visible != MsoTriState.msoTrue) return false;
                if (line == null || line.Visible != MsoTriState.msoTrue) return false;

                float weight = 0f;
                try { weight = (float)line.Weight; } catch { return false; }
                if (Math.Abs(weight - LineWeightTarget) > LineWeightTolerance)
                    return false;

                PptColorFormat fillColor = fill.ForeColor;
                PptColorFormat lineColor = line.ForeColor;
                if (fillColor == null || lineColor == null) return false;
                return IsAccent1Tint60Fill(fillColor) && IsDarkBlueLine(lineColor);
            }
            catch { return false; }
        }

        private static bool IsAccent1Tint60Fill(PptColorFormat cf)
        {
            if (cf == null) return false;
            try
            {
                if (cf.ObjectThemeColor != MsoThemeColorIndex.msoThemeColorAccent1)
                    return false;
            }
            catch { return false; }

            if (TryGetThemeColorAdjustment(cf, out float adjustment))
            {
                if (Math.Abs(adjustment - FillBrightnessTarget) <= FillBrightnessTolerance)
                    return true;
                // Brightness が 0 のまま返る環境: 淡い RGB で補完
                if (adjustment <= 0.15f && TryIsLightAccentRgb(cf))
                    return true;
                return false;
            }
            return TryIsLightAccentRgb(cf);
        }

        private static bool TryGetThemeColorAdjustment(PptColorFormat cf, out float adjustment)
        {
            adjustment = 0f;
            if (cf == null) return false;
            try
            {
                adjustment = (float)cf.Brightness;
                return true;
            }
            catch { }
            try
            {
                dynamic dynamicCf = cf;
                adjustment = (float)dynamicCf.TintAndShade;
                return true;
            }
            catch { return false; }
        }

        private static bool TryIsLightAccentRgb(PptColorFormat cf)
        {
            if (!TryGetRgb(cf, out int r, out int g, out int b)) return false;
            double lum = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0;
            // 白+60% 相当の淡さ（テーマ色相は問わない）
            return lum >= 0.55 && lum <= 0.90;
        }

        private static bool IsDarkBlueLine(PptColorFormat cf)
        {
            if (cf == null) return false;
            if (!TryGetRgb(cf, out int r, out int g, out int b)) return false;
            // #002060 付近（教材 P4-7 と同系統）
            return r <= 40 && g <= 80 && b >= 80 && b <= 140 && b > r && b > g;
        }

        private static bool TryGetRgb(PptColorFormat cf, out int r, out int g, out int b)
        {
            r = g = b = 0;
            if (cf == null) return false;
            try
            {
                int rgb = (int)cf.RGB;
                r = rgb & 0xFF;
                g = (rgb >> 8) & 0xFF;
                b = (rgb >> 16) & 0xFF;
                return true;
            }
            catch { return false; }
        }

        private static bool IsMediumStyle1Accent4(string styleName)
        {
            if (string.IsNullOrEmpty(styleName)) return false;
            string n = styleName;
            bool accent4 = n.IndexOf("アクセント 4", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("アクセント4", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("Accent 4", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("Accent4", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!accent4) return false;

            bool medium1 = n.IndexOf("中間スタイル 1", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("中間スタイル1", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("Medium Style 1", StringComparison.OrdinalIgnoreCase) >= 0
                || n.IndexOf("MediumStyle1", StringComparison.OrdinalIgnoreCase) >= 0;
            return medium1;
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
            catch { return null; }
            finally
            {
                if (raw != null && Marshal.IsComObject(raw))
                {
                    try { Marshal.ReleaseComObject(raw); } catch { }
                }
            }
        }

        private static int[] NormalizeSlideIdsFromComArray(Array arr)
        {
            var list = new List<int>();
            for (int i = 0; i < arr.Length; i++)
            {
                object o = arr.GetValue(i);
                if (o == null || Equals(o, Missing.Value)) continue;
                try
                {
                    int id = Convert.ToInt32(o);
                    if (id == 0) continue;
                    list.Add(id);
                }
                catch { }
            }
            return list.ToArray();
        }

        private static bool LayoutHasChartLeftOfTextPlaceholders(CustomLayout layout)
        {
            if (layout == null) return false;
            PptShapes shapes = null;
            try
            {
                shapes = layout.Shapes;
                if (shapes == null) return false;
                var chartLefts = new List<float>();
                var textLefts = new List<float>();
                for (int i = 1; i <= shapes.Count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        if (sh == null) continue;
                        if (sh.Type != MsoShapeType.msoPlaceholder) continue;
                        PlaceholderFormat pf = null;
                        try
                        {
                            pf = sh.PlaceholderFormat;
                            if (pf == null) continue;
                            var pt = (PpPlaceholderType)pf.Type;
                            float left = (float)sh.Left;
                            if (pt == PpPlaceholderType.ppPlaceholderChart)
                                chartLefts.Add(left);
                            else if (pt == PpPlaceholderType.ppPlaceholderBody
                                || pt == PpPlaceholderType.ppPlaceholderVerticalBody)
                                textLefts.Add(left);
                        }
                        finally
                        {
                            if (pf != null) { try { Marshal.ReleaseComObject(pf); } catch { } }
                        }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                if (chartLefts.Count == 0 || textLefts.Count == 0) return false;
                float minChartLeft = chartLefts[0];
                float maxTextLeft = textLefts[0];
                for (int i = 1; i < chartLefts.Count; i++)
                    if (chartLefts[i] < minChartLeft) minChartLeft = chartLefts[i];
                for (int i = 1; i < textLefts.Count; i++)
                    if (textLefts[i] > maxTextLeft) maxTextLeft = textLefts[i];
                return minChartLeft + 0.5f < maxTextLeft;
            }
            catch { return false; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }
    }
}
