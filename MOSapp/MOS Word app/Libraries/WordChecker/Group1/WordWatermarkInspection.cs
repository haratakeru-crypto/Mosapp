using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace Libraries.Group1
{
    /// <summary>
    /// 4-5 透かし判定。ギャラリー「サンプル２」は表示文字「サンプル」（ヘッダー内）。
    /// 「下書き1」は「下書き」＋斜め、「下書き2」は「下書き」＋横書き。
    /// WordChecker と VSTO ポーリングで同条件を共有する。
    /// </summary>
    public static class WordWatermarkInspection
    {
        private const int ContextCharsBefore = 2000;
        private const int ContextCharsAfter = 1500;

        private static readonly string[] ForbiddenWatermarkKeywords =
        {
            "社外秘", "至急", "案内", "転送厳禁", "CONFIDENTIAL", "URGENT", "DO NOT COPY"
        };

        private static readonly string[] DiagonalRotationLiteralMarkers =
        {
            "rotation:315", "rotation:-315", "rotation:-45", "rotation:45",
            "rot=\"315\"", "rot=\"-315\"", "rot=\"-45\"", "rot=\"45\""
        };

        private static readonly Regex RotationInStyleRegex = new Regex(
            @"rotation:\s*(-?\d+(?:\.\d+)?)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex HeaderBlockRegex = new Regex(
            @"<w:hdr\b[^>]*>.*?</w:hdr>",
            RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.IgnoreCase);

        public static string NormalizeXml(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return string.Empty;
            try
            {
                return xml.Normalize(NormalizationForm.FormKC);
            }
            catch
            {
                return xml;
            }
        }

        /// <summary>透かし「サンプル２」: ヘッダー内に「サンプル」または SAMPLE（下書き・禁止キーワードは除外）。</summary>
        public static bool IsSample2Watermark(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return false;
            if (HasForbiddenWatermark(normalizedXml))
                return false;
            if (normalizedXml.IndexOf("下書き", StringComparison.Ordinal) >= 0)
                return false;
            return ContainsSample2TextInWatermarkHeader(normalizedXml);
        }

        /// <summary>［デザイン］透かし「下書き1」: 文字「下書き」かつ斜め。社外秘・至急・下書き2（横）は false。</summary>
        public static bool IsDraft1Watermark(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return false;
            if (HasForbiddenWatermark(normalizedXml))
                return false;
            if (normalizedXml.IndexOf("下書き", StringComparison.Ordinal) < 0)
                return false;
            if (IsDraft2HorizontalWatermark(normalizedXml))
                return false;
            return ContainsDraft1DiagonalNearDraftText(normalizedXml);
        }

        /// <summary>透かし「下書き2」: 文字「下書き」かつ横書き（斜めでない）。</summary>
        public static bool IsDraft2HorizontalWatermark(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return false;
            if (normalizedXml.IndexOf("下書き", StringComparison.Ordinal) < 0)
                return false;

            foreach (int index in FindAllIndices(normalizedXml, "下書き"))
            {
                string context = GetContext(normalizedXml, index, ContextCharsBefore, ContextCharsAfter);
                if (ContextHasDiagonalRotation(context))
                    continue;
                if (ContextLooksHorizontalWatermark(context))
                    return true;
            }

            return false;
        }

        public static bool HasForbiddenWatermark(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return false;
            foreach (string keyword in ForbiddenWatermarkKeywords)
            {
                if (normalizedXml.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        /// <summary>スナップショット比較用の透かし指紋（タスク開始時と採点時の差分のみ。絶対判定しない）。</summary>
        public static string GetWatermarkFingerprint(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return "None";
            if (HasForbiddenWatermark(normalizedXml))
                return "Forbidden";
            if (IsDraft1Watermark(normalizedXml))
                return "Draft1Diagonal";
            if (IsDraft2HorizontalWatermark(normalizedXml))
                return "Draft2Horizontal";
            if (IsSample2Watermark(normalizedXml))
                return "Sample2";
            if (normalizedXml.IndexOf("下書き", StringComparison.Ordinal) >= 0)
                return "DraftOther";
            return "None";
        }

        public static string GetWatermarkFingerprintFromDocument(Document doc)
        {
            if (doc == null)
                return "None";
            try
            {
                return GetWatermarkFingerprint(NormalizeXml(doc.WordOpenXML));
            }
            catch
            {
                return "None";
            }
        }

        /// <summary>先頭セクションのページ上辺・下辺罫線（4-6 採点と同一形式）。</summary>
        public static string GetPageBorderFingerprint(Document doc)
        {
            if (doc == null)
                return string.Empty;
            Section sec = null;
            Borders borders = null;
            Border top = null;
            Border bottom = null;
            try
            {
                if (doc.Sections.Count < 1)
                    return string.Empty;
                sec = doc.Sections[1];
                borders = sec.Borders;
                top = borders[WdBorderType.wdBorderTop];
                bottom = borders[WdBorderType.wdBorderBottom];
                return string.Format(CultureInfo.InvariantCulture,
                    "{0},{1},{2},{3}",
                    (int)top.LineStyle, (int)top.LineWidth, (int)bottom.LineStyle, (int)bottom.LineWidth);
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                if (bottom != null)
                    Marshal.ReleaseComObject(bottom);
                if (top != null)
                    Marshal.ReleaseComObject(top);
                if (borders != null)
                    Marshal.ReleaseComObject(borders);
                if (sec != null)
                    Marshal.ReleaseComObject(sec);
            }
        }

        /// <summary>4-4: スタイルセット「線（シンプル）」— 見出し1に0.5pt単線、見出し2に下罫線なし。</summary>
        public static bool IsDocumentStyleSetLineSimple(Document doc)
        {
            if (doc == null)
                return false;
            if (!TryGetStyleBottomBorder(doc, WdBuiltinStyle.wdStyleHeading1, out int h1Style, out int h1Width))
                return false;
            if (!TryGetStyleBottomBorder(doc, WdBuiltinStyle.wdStyleHeading2, out int h2Style, out int _))
                return false;
            bool h1Ok = h1Style == (int)WdLineStyle.wdLineStyleSingle
                        && h1Width == (int)WdLineWidth.wdLineWidth050pt;
            bool h2NoBorder = h2Style == 0 || h2Style == (int)WdLineStyle.wdLineStyleNone;
            return h1Ok && h2NoBorder;
        }

        /// <summary>4-4 否定用: スタイルセット「線（スタイリッシュ）」— 見出し1に0.5pt単線かつ見出し2にも下罫線あり。</summary>
        public static bool IsDocumentStyleSetLineStylish(Document doc)
        {
            if (doc == null)
                return false;
            if (!TryGetStyleBottomBorder(doc, WdBuiltinStyle.wdStyleHeading1, out int h1Style, out int h1Width))
                return false;
            if (!TryGetStyleBottomBorder(doc, WdBuiltinStyle.wdStyleHeading2, out int h2Style, out int h2Width))
                return false;
            bool h1Ok = h1Style == (int)WdLineStyle.wdLineStyleSingle
                        && h1Width == (int)WdLineWidth.wdLineWidth050pt;
            bool h2HasBorder = h2Style == (int)WdLineStyle.wdLineStyleSingle && h2Width > 0;
            return h1Ok && h2HasBorder;
        }

        /// <summary>4-6: ページ罫線4辺が純粋な accent3（themeTint/themeShade なし）か。</summary>
        public static bool ArePageBordersPureAccent3(Document document)
        {
            if (document == null)
                return false;
            try
            {
                string docXml = null;
                try { docXml = document.WordOpenXML; } catch { }
                if (!string.IsNullOrEmpty(docXml) && TryArePgBordersPureAccent3(docXml))
                    return true;
                return ArePageBordersPureAccent3Com(document);
            }
            catch
            {
                return false;
            }
        }

        private static bool TryArePgBordersPureAccent3(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return false;
            int start = xml.IndexOf("<w:pgBorders", StringComparison.OrdinalIgnoreCase);
            if (start < 0)
                return false;
            int end = xml.IndexOf("</w:pgBorders>", start, StringComparison.OrdinalIgnoreCase);
            if (end < 0)
                return false;
            end += "</w:pgBorders>".Length;
            string pgBordersXml = xml.Substring(start, end - start);
            return IsPgBorderSidePureAccent3(pgBordersXml, "top")
                && IsPgBorderSidePureAccent3(pgBordersXml, "left")
                && IsPgBorderSidePureAccent3(pgBordersXml, "bottom")
                && IsPgBorderSidePureAccent3(pgBordersXml, "right");
        }

        private static bool IsPgBorderSidePureAccent3(string pgBordersXml, string sideName)
        {
            var match = Regex.Match(pgBordersXml, $@"<w:{sideName}\s+([^/>]*)/>", RegexOptions.IgnoreCase);
            if (!match.Success)
                return false;
            string attrs = match.Groups[1].Value;
            if (!Regex.IsMatch(attrs, @"w:themeColor\s*=\s*""accent3""", RegexOptions.IgnoreCase))
                return false;
            if (HasNonPureThemeTintOrShade(attrs))
                return false;
            return true;
        }

        private static bool HasNonPureThemeTintOrShade(string attrs)
        {
            Match mTint = Regex.Match(attrs, @"w:themeTint\s*=\s*""([0-9A-Fa-f]+)""", RegexOptions.IgnoreCase);
            if (mTint.Success)
            {
                int v = Convert.ToInt32(mTint.Groups[1].Value, 16);
                if (v != 0)
                    return true;
            }
            Match mShade = Regex.Match(attrs, @"w:themeShade\s*=\s*""([0-9A-Fa-f]+)""", RegexOptions.IgnoreCase);
            if (mShade.Success)
            {
                int v = Convert.ToInt32(mShade.Groups[1].Value, 16);
                if (v != 0)
                    return true;
            }
            return false;
        }

        private static bool ArePageBordersPureAccent3Com(Document document)
        {
            Section sec = null;
            Borders borders = null;
            try
            {
                if (document.Sections.Count < 1)
                    return false;
                sec = document.Sections[1];
                borders = sec.Borders;
                foreach (WdBorderType side in new[] { WdBorderType.wdBorderTop, WdBorderType.wdBorderBottom, WdBorderType.wdBorderLeft, WdBorderType.wdBorderRight })
                {
                    Border b = null;
                    try
                    {
                        b = borders[side];
                        if (b == null)
                            return false;
                        dynamic color = b.Color;
                        WdThemeColorIndex theme = WdThemeColorIndex.wdNotThemeColor;
                        try { theme = (WdThemeColorIndex)color.ObjectThemeColor; } catch { return false; }
                        if (theme != WdThemeColorIndex.wdThemeColorAccent3)
                            return false;
                        float tint = 0f;
                        try { tint = (float)color.TintAndShade; } catch { }
                        if (Math.Abs(tint) > 0.05f)
                            return false;
                    }
                    finally
                    {
                        if (b != null)
                            Marshal.ReleaseComObject(b);
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (borders != null)
                    Marshal.ReleaseComObject(borders);
                if (sec != null)
                    Marshal.ReleaseComObject(sec);
            }
        }

        private static bool TryGetStyleBottomBorder(Document doc, WdBuiltinStyle styleId, out int lineStyle, out int lineWidth)
        {
            lineStyle = 0;
            lineWidth = 0;
            Style style = null;
            Borders borders = null;
            Border bottom = null;
            try
            {
                style = doc.Styles[styleId];
                if (style == null)
                    return false;
                borders = style.ParagraphFormat.Borders;
                bottom = borders[WdBorderType.wdBorderBottom];
                if (bottom == null)
                    return true;
                try { lineStyle = (int)bottom.LineStyle; } catch { }
                try { lineWidth = (int)bottom.LineWidth; } catch { }
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (bottom != null)
                    Marshal.ReleaseComObject(bottom);
                if (borders != null)
                    Marshal.ReleaseComObject(borders);
                if (style != null)
                    Marshal.ReleaseComObject(style);
            }
        }

        /// <summary>4-7 後: 禁止透かし・下書き1/2 相当の透かしが残っていない。</summary>
        public static bool IsWatermarkClearedForTask47FollowUp(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return true;
            if (HasForbiddenWatermark(normalizedXml))
                return false;
            if (IsDraft1Watermark(normalizedXml))
                return false;
            if (IsDraft2HorizontalWatermark(normalizedXml))
                return false;
            if (IsSample2Watermark(normalizedXml))
                return false;
            return normalizedXml.IndexOf("下書き", StringComparison.Ordinal) < 0
                && normalizedXml.IndexOf("サンプル", StringComparison.Ordinal) < 0;
        }

        private static bool ContainsSample2TextInWatermarkHeader(string normalizedXml)
        {
            foreach (Match match in HeaderBlockRegex.Matches(normalizedXml))
            {
                string hdr = match.Value;
                if (hdr.IndexOf("サンプル", StringComparison.Ordinal) >= 0)
                    return true;
                if (hdr.IndexOf("SAMPLE", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            return false;
        }

        private static bool ContainsDraft1DiagonalNearDraftText(string normalizedXml)
        {
            foreach (int index in FindAllIndices(normalizedXml, "下書き"))
            {
                string context = GetContext(normalizedXml, index, ContextCharsBefore, ContextCharsAfter);
                if (ContextHasDiagonalRotation(context))
                    return true;
            }
            return false;
        }

        private static bool ContextHasDiagonalRotation(string context)
        {
            if (string.IsNullOrEmpty(context))
                return false;

            foreach (string marker in DiagonalRotationLiteralMarkers)
            {
                if (context.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            foreach (Match match in RotationInStyleRegex.Matches(context))
            {
                if (!TryParseRotationDegrees(match.Groups[1].Value, out double degrees))
                    continue;
                if (IsDiagonalRotationDegrees(degrees))
                    return true;
            }

            return false;
        }

        private static bool ContextLooksHorizontalWatermark(string context)
        {
            if (string.IsNullOrEmpty(context))
                return false;

            if (context.IndexOf("rotation:0", StringComparison.OrdinalIgnoreCase) >= 0
                || context.IndexOf("rotation:0deg", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            foreach (Match match in RotationInStyleRegex.Matches(context))
            {
                if (TryParseRotationDegrees(match.Groups[1].Value, out double degrees)
                    && Math.Abs(degrees) < 2)
                    return true;
            }

            // 横書きプリセットは textDirection が lrTb で rotation 指定が無いことが多い
            if (context.IndexOf("lrTb", StringComparison.OrdinalIgnoreCase) >= 0
                && context.IndexOf("下書き", StringComparison.Ordinal) >= 0)
                return true;

            return false;
        }

        private static bool TryParseRotationDegrees(string value, out double degrees)
        {
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out degrees);
        }

        /// <summary>Word 透かしの斜め（下書き1）: 0° 付近以外の典型角。</summary>
        private static bool IsDiagonalRotationDegrees(double degrees)
        {
            degrees = NormalizeDegrees(degrees);
            if (degrees < 2)
                return false;
            // 45°, 135°, 225°, 315° 付近（±8°）
            double[] diagonalAngles = { 45, 135, 225, 315 };
            foreach (double target in diagonalAngles)
            {
                double diff = Math.Abs(degrees - target);
                if (diff <= 8 || Math.Abs(diff - 360) <= 8)
                    return true;
            }
            return false;
        }

        private static double NormalizeDegrees(double degrees)
        {
            degrees %= 360;
            if (degrees < 0)
                degrees += 360;
            return degrees;
        }

        private static string GetContext(string text, int anchorIndex, int charsBefore, int charsAfter)
        {
            int start = Math.Max(0, anchorIndex - charsBefore);
            int end = Math.Min(text.Length, anchorIndex + charsAfter);
            return text.Substring(start, end - start);
        }

        private static IEnumerable<int> FindAllIndices(string text, string value)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(value))
                yield break;

            int index = 0;
            while (index < text.Length)
            {
                index = text.IndexOf(value, index, StringComparison.Ordinal);
                if (index < 0)
                    yield break;
                yield return index;
                index += value.Length;
            }
        }
    }
}
