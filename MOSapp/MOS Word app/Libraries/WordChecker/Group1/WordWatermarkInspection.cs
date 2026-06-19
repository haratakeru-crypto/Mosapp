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

        private static readonly string[] StyleSetLineStyleIdsHeading2 =
        {
            "Heading2", "heading2", "見出し2"
        };

        private static readonly string[] StyleSetLineStyleIdsTitle =
        {
            "Title", "title", "表題"
        };

        private static readonly string[] StyleSetLineStyleIdsHeading1 =
        {
            "Heading1", "heading1", "見出し1"
        };

        /// <summary>4-4: 見出し1/2・表題スタイル定義の OpenXML スナップショット（スタイルセット切替検知用）。</summary>
        public static string GetStyleSetLineFingerprint(Document doc)
        {
            var parts = new List<string>(10);
            AppendStyleSetLineComMetrics(parts, doc);

            string stylesXml = TryGetOpenXmlStylesBlob(doc);
            if (!string.IsNullOrEmpty(stylesXml))
            {
                AppendStyleBlockFingerprint(parts, stylesXml, StyleSetLineStyleIdsHeading1);
                AppendStyleBlockFingerprint(parts, stylesXml, StyleSetLineStyleIdsHeading2);
                AppendStyleBlockFingerprint(parts, stylesXml, StyleSetLineStyleIdsTitle);
            }

            parts.Add("H2ParaBdr:" + (HasAnyHeading2ParagraphBottomBorderCom(doc) ? "1" : "0"));
            parts.Add("H2XmlBdr:" + (HasHeading2ParagraphBottomBorderInOpenXml(doc) ? "1" : "0"));
            return string.Join("|", parts);
        }

        private static void AppendStyleSetLineComMetrics(List<string> parts, Document doc)
        {
            if (doc == null)
                return;

            if (TryGetStyleFontMetrics(doc, WdBuiltinStyle.wdStyleHeading1, out _, out float h1Size, out int h1Theme, out _))
            {
                parts.Add(string.Format(CultureInfo.InvariantCulture, "H1Sz={0}", h1Size));
                parts.Add("H1Th=" + h1Theme);
            }

            if (TryGetStyleFontMetrics(doc, WdBuiltinStyle.wdStyleHeading2, out _, out float h2Size, out int h2Theme, out _))
            {
                parts.Add(string.Format(CultureInfo.InvariantCulture, "H2Sz={0}", h2Size));
                parts.Add("H2Th=" + h2Theme);
            }

            if (TryGetStyleBottomBorderColor(doc, WdBuiltinStyle.wdStyleHeading1, out int h1BorderColor))
                parts.Add("H1BdrClr=" + h1BorderColor);
        }

        private static bool TryGetStyleBottomBorderColor(Document doc, WdBuiltinStyle styleId, out int color)
        {
            color = 0;
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
                    return false;
                try
                {
                    color = (int)bottom.Color;
                    return true;
                }
                catch
                {
                    return false;
                }
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

        private static void AppendStyleBlockFingerprint(List<string> parts, string stylesXml, string[] styleIds)
        {
            foreach (string styleId in styleIds)
            {
                if (string.IsNullOrEmpty(styleId))
                    continue;
                string pattern = $@"<w:style\b[^>]*\bw:styleId=""{Regex.Escape(styleId)}""[^>]*>([\s\S]*?)</w:style>";
                Match styleMatch = Regex.Match(stylesXml, pattern, RegexOptions.IgnoreCase);
                if (!styleMatch.Success)
                    continue;
                parts.Add(styleId + ":" + styleMatch.Groups[1].Value);
                return;
            }
        }

        /// <summary>4-4: スタイルセット「線（シンプル）」— 見出し1に0.5pt単線、見出し2・表題に下罫線なし、見出し1/2フォントプロファイル一致。</summary>
        public static bool IsDocumentStyleSetLineSimple(Document doc)
        {
            if (doc == null || IsDocumentStyleSetLineStylish(doc))
                return false;

            if (!TryGetStyleBottomBorder(doc, WdBuiltinStyle.wdStyleHeading1, out int h1Style, out int h1Width))
                return false;
            if (!TryGetStyleBottomBorder(doc, WdBuiltinStyle.wdStyleHeading2, out int h2Style, out int h2Width))
                return false;

            bool h1Ok = h1Style == (int)WdLineStyle.wdLineStyleSingle
                        && h1Width == (int)WdLineWidth.wdLineWidth050pt;
            if (!h1Ok)
                return false;

            if (HasMeaningfulBottomBorder(h2Style, h2Width))
                return false;
            if (HasStyleBottomBorderCom(doc, WdBuiltinStyle.wdStyleTitle))
                return false;
            if (HasAnyHeading2ParagraphBottomBorder(doc))
                return false;

            string stylesXml = TryGetOpenXmlStylesBlob(doc);
            if (StyleBlockHasBottomBorder(stylesXml, StyleSetLineStyleIdsHeading2))
                return false;
            if (StyleBlockHasBottomBorder(stylesXml, StyleSetLineStyleIdsTitle))
                return false;

            return HasLineSimpleFontProfile(doc);
        }

        /// <summary>4-4 否定用: スタイルセット「線（スタイリッシュ）」等、シンプル以外の線スタイルセット。</summary>
        public static bool IsDocumentStyleSetLineStylish(Document doc)
        {
            if (doc == null)
                return false;

            if (HasStyleSetLineStylishAppearance(doc))
                return true;

            if (HasAnyHeading2ParagraphBottomBorder(doc))
                return true;

            if (HasStyleSetLineStylishByHeading2Com(doc))
                return true;
            if (HasStyleBottomBorderCom(doc, WdBuiltinStyle.wdStyleTitle))
                return true;

            string stylesXml = TryGetOpenXmlStylesBlob(doc);
            if (StyleBlockHasBottomBorder(stylesXml, StyleSetLineStyleIdsHeading2))
                return true;
            if (StyleBlockHasBottomBorder(stylesXml, StyleSetLineStyleIdsTitle))
                return true;
            if (IsHeading1BottomBorderAccentThemed(doc, stylesXml))
                return true;

            return CountHeadingLikeStylesWithBottomBorder(stylesXml) >= 2;
        }

        /// <summary>本文中の見出し2段落に実効下罫線があるか（COM + document.xml）。</summary>
        private static bool HasAnyHeading2ParagraphBottomBorder(Document doc)
        {
            if (HasAnyHeading2ParagraphBottomBorderCom(doc))
                return true;
            return HasHeading2ParagraphBottomBorderInOpenXml(doc);
        }

        private static bool HasAnyHeading2ParagraphBottomBorderCom(Document doc)
        {
            if (doc == null)
                return false;

            Paragraphs paragraphs = null;
            try
            {
                paragraphs = doc.Paragraphs;
                int count = 0;
                try { count = paragraphs.Count; } catch { return false; }

                for (int i = 1; i <= count; i++)
                {
                    Paragraph para = null;
                    try
                    {
                        para = paragraphs[i];
                        if (!IsHeading2Paragraph(para))
                            continue;
                        if (TryGetParagraphEffectiveBottomBorder(para, out int lineStyle, out int lineWidth)
                            && HasMeaningfulBottomBorder(lineStyle, lineWidth))
                            return true;
                    }
                    catch { }
                    finally
                    {
                        if (para != null)
                            Marshal.ReleaseComObject(para);
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (paragraphs != null)
                    Marshal.ReleaseComObject(paragraphs);
            }
        }

        private static bool HasHeading2ParagraphBottomBorderInOpenXml(Document doc)
        {
            string xml = TryGetFullWordOpenXml(doc);
            if (string.IsNullOrEmpty(xml))
                return false;

            string bodyXml = ExtractOpenXmlDocumentBodyBlob(xml);
            if (string.IsNullOrEmpty(bodyXml))
                bodyXml = xml;

            foreach (Match pMatch in Regex.Matches(bodyXml, @"<w:p\b[^>]*>([\s\S]*?)</w:p>", RegexOptions.IgnoreCase))
            {
                string pInner = pMatch.Groups[1].Value;
                if (!ParagraphBlockUsesHeading2Style(pInner))
                    continue;
                if (ParagraphPropertiesBlockHasBottomBorder(pInner))
                    return true;
            }

            return false;
        }

        private static bool IsHeading2Paragraph(Paragraph para)
        {
            if (para == null)
                return false;

            try
            {
                if (IsHeading2StyleName(GetStyleNameFromObject(para.get_Style())))
                    return true;

                Range range = para.Range;
                if (range != null)
                {
                    try
                    {
                        if (range.Characters.Count > 0)
                        {
                            Range firstChar = range.Characters[1];
                            try
                            {
                                if (IsHeading2StyleName(GetStyleNameFromObject(firstChar.get_Style())))
                                    return true;
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(firstChar);
                            }
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(range);
                    }
                }
            }
            catch { }

            return false;
        }

        private static bool IsHeading2StyleName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return false;

            string normalized = name.Replace(" ", string.Empty).Replace("　", string.Empty);
            return normalized.IndexOf("Heading2", StringComparison.OrdinalIgnoreCase) >= 0
                || normalized.IndexOf("heading2", StringComparison.OrdinalIgnoreCase) >= 0
                || normalized.IndexOf("見出し2", StringComparison.Ordinal) >= 0
                || normalized.IndexOf("見出し２", StringComparison.Ordinal) >= 0;
        }

        private static string GetStyleNameFromObject(object styleObj)
        {
            if (styleObj == null)
                return string.Empty;

            if (styleObj is string styleName)
                return styleName ?? string.Empty;

            if (styleObj is Style style)
            {
                try { return style.NameLocal ?? string.Empty; }
                catch { }
            }

            return styleObj.ToString() ?? string.Empty;
        }

        private static bool TryGetParagraphEffectiveBottomBorder(Paragraph para, out int lineStyle, out int lineWidth)
        {
            lineStyle = 0;
            lineWidth = 0;
            if (para == null)
                return false;

            Borders borders = null;
            Border bottom = null;
            Range range = null;
            try
            {
                range = para.Range;
                borders = range.ParagraphFormat.Borders;
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
                if (range != null)
                    Marshal.ReleaseComObject(range);
            }
        }

        private static string TryGetFullWordOpenXml(Document doc)
        {
            if (doc == null)
                return string.Empty;
            try
            {
                return doc.WordOpenXML ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ExtractOpenXmlDocumentBodyBlob(string wordOpenXml)
        {
            if (string.IsNullOrEmpty(wordOpenXml))
                return string.Empty;

            var sb = new StringBuilder(wordOpenXml.Length);
            foreach (Match m in Regex.Matches(
                wordOpenXml,
                @"<pkg:xmlData[^>]*>([\s\S]*?)</pkg:xmlData>",
                RegexOptions.IgnoreCase))
            {
                string chunk = m.Groups[1].Value;
                if (chunk.IndexOf("<w:body", StringComparison.OrdinalIgnoreCase) >= 0
                    || chunk.IndexOf("<w:p ", StringComparison.OrdinalIgnoreCase) >= 0
                    || chunk.IndexOf("<w:p>", StringComparison.OrdinalIgnoreCase) >= 0)
                    sb.Append('\n').Append(chunk);
            }

            if (sb.Length > 0)
                return sb.ToString();
            return wordOpenXml;
        }

        private static bool ParagraphBlockUsesHeading2Style(string paragraphInnerXml)
        {
            if (string.IsNullOrEmpty(paragraphInnerXml))
                return false;

            Match styleMatch = Regex.Match(
                paragraphInnerXml,
                @"<w:pStyle\b[^>]*\bw:val=""([^""]+)""",
                RegexOptions.IgnoreCase);
            if (!styleMatch.Success)
                return false;

            string styleVal = styleMatch.Groups[1].Value ?? string.Empty;
            foreach (string id in StyleSetLineStyleIdsHeading2)
            {
                if (string.Equals(styleVal, id, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return string.Equals(styleVal, "2", StringComparison.Ordinal);
        }

        private static bool ParagraphPropertiesBlockHasBottomBorder(string paragraphInnerXml)
        {
            if (string.IsNullOrEmpty(paragraphInnerXml))
                return false;

            Match pPrMatch = Regex.Match(
                paragraphInnerXml,
                @"<w:pPr\b[^>]*>([\s\S]*?)</w:pPr>",
                RegexOptions.IgnoreCase);
            string pPrBlock = pPrMatch.Success ? pPrMatch.Groups[1].Value : paragraphInnerXml;
            return TryGetParagraphBottomBorderFromStyleBlock(pPrBlock, out bool hasBorder) && hasBorder;
        }

        private static bool HasStyleSetLineStylishByHeading2Com(Document doc)
        {
            if (!TryGetStyleBottomBorder(doc, WdBuiltinStyle.wdStyleHeading1, out int h1Style, out int h1Width))
                return false;
            if (!TryGetStyleBottomBorder(doc, WdBuiltinStyle.wdStyleHeading2, out int h2Style, out int h2Width))
                return false;
            bool h1Ok = h1Style == (int)WdLineStyle.wdLineStyleSingle
                        && h1Width == (int)WdLineWidth.wdLineWidth050pt;
            return h1Ok && HasMeaningfulBottomBorder(h2Style, h2Width);
        }

        /// <summary>
        /// 4-4: 線（スタイリッシュ）の見た目 — 見出し2サイズ拡大・見出し1サイズ拡大・アクセント2色など。
        /// Word 365 実測: シンプル H1/H2=18/14pt theme=4、スタイリッシュ H1/H2=20/18pt theme=13/5。
        /// </summary>
        private static bool HasStyleSetLineStylishAppearance(Document doc)
        {
            if (!TryGetStyleFontMetrics(doc, WdBuiltinStyle.wdStyleHeading2, out _, out float h2Size, out int h2Theme, out _))
                return false;
            if (!TryGetStyleFontMetrics(doc, WdBuiltinStyle.wdStyleHeading1, out _, out float h1Size, out int h1Theme, out _))
                return false;

            if (h2Size >= 17f)
                return true;
            if (h1Size >= 19f)
                return true;
            if (h2Theme == (int)WdThemeColorIndex.wdThemeColorAccent2)
                return true;
            if (h1Theme == 13)
                return true;

            return false;
        }

        /// <summary>4-4: 線（シンプル）の見出し1/2スタイル定義フォントプロファイル。</summary>
        private static bool HasLineSimpleFontProfile(Document doc)
        {
            if (!TryGetStyleFontMetrics(doc, WdBuiltinStyle.wdStyleHeading2, out _, out float h2Size, out int h2Theme, out _))
                return false;
            if (!TryGetStyleFontMetrics(doc, WdBuiltinStyle.wdStyleHeading1, out _, out float h1Size, out int h1Theme, out _))
                return false;

            if (h2Size > 15f)
                return false;
            if (h1Size > 18.5f)
                return false;

            int accent1 = (int)WdThemeColorIndex.wdThemeColorAccent1;
            int notTheme = (int)WdThemeColorIndex.wdNotThemeColor;
            if (h2Theme != notTheme && h2Theme != accent1)
                return false;
            if (h1Theme != notTheme && h1Theme != accent1)
                return false;

            return true;
        }

        private static bool TryGetStyleFontMetrics(
            Document doc,
            WdBuiltinStyle styleId,
            out int bold,
            out float size,
            out int theme,
            out float tint)
        {
            bold = 0;
            size = 0f;
            theme = (int)WdThemeColorIndex.wdNotThemeColor;
            tint = float.NaN;

            Style style = null;
            Font font = null;
            try
            {
                style = doc.Styles[styleId];
                if (style == null)
                    return false;
                font = style.Font;
                if (font == null)
                    return false;
                try { bold = font.Bold; } catch { }
                try { size = (float)font.Size; } catch { }
                try { theme = (int)font.TextColor.ObjectThemeColor; } catch { }
                try { tint = font.TextColor.TintAndShade; } catch { }
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (font != null)
                    Marshal.ReleaseComObject(font);
                if (style != null)
                    Marshal.ReleaseComObject(style);
            }
        }

        private static bool HasStyleBottomBorderCom(Document doc, WdBuiltinStyle styleId)
        {
            if (!TryGetStyleBottomBorder(doc, styleId, out int lineStyle, out int lineWidth))
                return false;
            return HasMeaningfulBottomBorder(lineStyle, lineWidth);
        }

        private static bool HasMeaningfulBottomBorder(int lineStyle, int lineWidth)
        {
            if (lineStyle == 0 || lineStyle == (int)WdLineStyle.wdLineStyleNone)
                return false;
            return lineWidth > 0 || lineStyle != (int)WdLineStyle.wdLineStyleNone;
        }

        /// <summary>見出し1下罫線がアクセント系テーマ色なら「線（スタイリッシュ）」側とみなす（シンプルは text/dark 系が多い）。</summary>
        private static bool IsHeading1BottomBorderAccentThemed(Document doc, string stylesXml)
        {
            if (TryGetStyleBottomBorderThemeFromOpenXml(stylesXml, StyleSetLineStyleIdsHeading1, out string theme))
                return IsAccentThemeColorName(theme);
            return false;
        }

        private static bool IsAccentThemeColorName(string theme)
        {
            if (string.IsNullOrEmpty(theme))
                return false;
            return theme.IndexOf("accent", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string TryGetOpenXmlStylesBlob(Document doc)
        {
            if (doc == null)
                return string.Empty;
            try
            {
                string xml = doc.WordOpenXML;
                if (string.IsNullOrEmpty(xml))
                    return string.Empty;
                return ExtractOpenXmlStylesBlob(xml);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ExtractOpenXmlStylesBlob(string wordOpenXml)
        {
            if (string.IsNullOrEmpty(wordOpenXml))
                return string.Empty;

            var sb = new StringBuilder(wordOpenXml.Length);
            foreach (Match m in Regex.Matches(
                wordOpenXml,
                @"<pkg:xmlData[^>]*>([\s\S]*?)</pkg:xmlData>",
                RegexOptions.IgnoreCase))
            {
                string chunk = m.Groups[1].Value;
                if (chunk.IndexOf("<w:style", StringComparison.OrdinalIgnoreCase) >= 0
                    || chunk.IndexOf("<w:styles", StringComparison.OrdinalIgnoreCase) >= 0)
                    sb.Append('\n').Append(chunk);
            }

            if (sb.Length > 0)
                return sb.ToString();
            return wordOpenXml;
        }

        private static bool StyleBlockHasBottomBorder(string stylesXml, params string[] styleIds)
        {
            if (string.IsNullOrEmpty(stylesXml) || styleIds == null)
                return false;
            foreach (string styleId in styleIds)
            {
                if (string.IsNullOrEmpty(styleId))
                    continue;
                if (TryGetStyleBottomBorderFromOpenXml(stylesXml, styleId, out bool hasBorder) && hasBorder)
                    return true;
            }
            return false;
        }

        private static bool TryGetStyleBottomBorderFromOpenXml(string stylesXml, string styleId, out bool hasBorder)
        {
            hasBorder = false;
            if (string.IsNullOrEmpty(stylesXml) || string.IsNullOrEmpty(styleId))
                return false;

            string pattern = $@"<w:style\b[^>]*\bw:styleId=""{Regex.Escape(styleId)}""[^>]*>([\s\S]*?)</w:style>";
            Match styleMatch = Regex.Match(stylesXml, pattern, RegexOptions.IgnoreCase);
            if (!styleMatch.Success)
                return false;

            return TryGetParagraphBottomBorderFromStyleBlock(styleMatch.Groups[1].Value, out hasBorder);
        }

        private static bool TryGetStyleBottomBorderThemeFromOpenXml(string stylesXml, string[] styleIds, out string themeColor)
        {
            themeColor = null;
            if (string.IsNullOrEmpty(stylesXml) || styleIds == null)
                return false;

            foreach (string styleId in styleIds)
            {
                if (string.IsNullOrEmpty(styleId))
                    continue;
                string pattern = $@"<w:style\b[^>]*\bw:styleId=""{Regex.Escape(styleId)}""[^>]*>([\s\S]*?)</w:style>";
                Match styleMatch = Regex.Match(stylesXml, pattern, RegexOptions.IgnoreCase);
                if (!styleMatch.Success)
                    continue;
                if (!TryGetParagraphBottomBorderFromStyleBlock(styleMatch.Groups[1].Value, out bool hasBorder) || !hasBorder)
                    continue;

                Match bottomMatch = Regex.Match(
                    styleMatch.Groups[1].Value,
                    @"<w:bottom\b([^/>]*)/>",
                    RegexOptions.IgnoreCase);
                if (!bottomMatch.Success)
                    continue;

                Match themeMatch = Regex.Match(bottomMatch.Groups[1].Value, @"w:themeColor=""([^""]+)""", RegexOptions.IgnoreCase);
                if (themeMatch.Success)
                {
                    themeColor = themeMatch.Groups[1].Value;
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetParagraphBottomBorderFromStyleBlock(string styleBlock, out bool hasBorder)
        {
            hasBorder = false;
            if (string.IsNullOrEmpty(styleBlock))
                return false;

            Match bottomMatch = Regex.Match(styleBlock, @"<w:bottom\b([^/>]*)/>", RegexOptions.IgnoreCase);
            if (!bottomMatch.Success)
                return false;

            Match valMatch = Regex.Match(bottomMatch.Groups[1].Value, @"w:val=""([^""]+)""", RegexOptions.IgnoreCase);
            if (!valMatch.Success)
            {
                hasBorder = true;
                return true;
            }

            string val = valMatch.Groups[1].Value.Trim().ToLowerInvariant();
            hasBorder = val != "none" && val != "nil" && val != "hidden";
            return true;
        }

        private static int CountHeadingLikeStylesWithBottomBorder(string stylesXml)
        {
            if (string.IsNullOrEmpty(stylesXml))
                return 0;

            int count = 0;
            foreach (Match styleMatch in Regex.Matches(stylesXml, @"<w:style\b[^>]*>([\s\S]*?)</w:style>", RegexOptions.IgnoreCase))
            {
                string fullBlock = styleMatch.Value;
                if (!Regex.IsMatch(fullBlock, @"w:styleId=""(Heading\d|Title|表題|見出し)", RegexOptions.IgnoreCase))
                    continue;
                if (TryGetParagraphBottomBorderFromStyleBlock(styleMatch.Groups[1].Value, out bool hasBorder) && hasBorder)
                    count++;
            }
            return count;
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
