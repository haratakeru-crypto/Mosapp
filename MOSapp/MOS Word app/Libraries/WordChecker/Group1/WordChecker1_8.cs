using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_8
    {
        private const string FootnoteTargetText803 = "電子機器に悪影響を与えるプログラム";

        public bool CheckTask_1_8_01() => RunCheck(CheckTask_1_8_01_Impl);
        public bool CheckTask_1_8_02() => RunCheck(CheckTask_1_8_02_Impl);
        public bool CheckTask_1_8_03() => RunCheck(CheckTask_1_8_03_Impl);
        public bool CheckTask_1_8_04() => RunCheck(CheckTask_1_8_04_Impl);
        public bool CheckTask_1_8_05() => RunCheck(CheckTask_1_8_05_Impl);
        public bool CheckTask_1_8_06() => RunCheck(CheckTask_1_8_06_Impl);
        public bool CheckTask_1_8_07() => RunCheck(CheckTask_1_8_07_Impl);

        private bool RunCheck(Func<string, bool> task)
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(WordScoreExplanation.UnavailableText);
                return task(filePath);
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
        }

        private bool CheckTask_1_8_01_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                // 日本語ギャラリー: 自動作成の目次2 = 見出し「目次」、目次1 = 見出し「内容」
                var txbxMatches = Regex.Matches(xml, @"<w:txbxContent\b[^>]*>.*?</w:txbxContent>", RegexOptions.Singleline);
                bool foundToc1 = false;

                foreach (Match match in txbxMatches)
                {
                    string shapeXml = match.Value;
                    bool hasTOCField = Regex.IsMatch(shapeXml, @"<w:instrText\b[^>]*>\s*TOC\b", RegexOptions.IgnoreCase);
                    if (!hasTOCField)
                        continue;

                    bool hasMokujiTitle = Regex.IsMatch(shapeXml, @"<w:t\b[^>]*>目次</w:t>");
                    bool hasNaiyouTitle = Regex.IsMatch(shapeXml, @"<w:t\b[^>]*>内容</w:t>");

                    if (hasMokujiTitle)
                        return true;
                    if (hasNaiyouTitle)
                        foundToc1 = true;
                }

                if (foundToc1)
                    return Miss("図形に「自動作成の目次1」が設定されています。");
                return Miss("図形に「自動作成の目次2」の目次が挿入されていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_02_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                const string expectedFootnote = FootnoteTargetText803;
                const string missNotOnMalware = "マルウェアに「電子機器に悪影響を与えるプログラム」の脚注が設定されていません。";

                var documentPartMatch = Regex.Match(xml, @"<pkg:part pkg:name=""/word/document\.xml""[^>]*>.*?</pkg:part>", RegexOptions.Singleline);
                string bodyXml = documentPartMatch.Success ? documentPartMatch.Value : xml;
                var paragraphs = Regex.Matches(bodyXml, @"<w:p\b[^>]*>.*?</w:p>", RegexOptions.Singleline);

                string malwareParagraphXml = null;
                foreach (Match p in paragraphs)
                {
                    string pTextOnly = Regex.Replace(p.Value, @"<[^>]+>", "");
                    if (!pTextOnly.Contains("マルウェア"))
                        continue;
                    malwareParagraphXml = p.Value;
                    break;
                }

                if (malwareParagraphXml == null)
                    return Miss("「マルウェア」が見つかりません。");

                // 文末脚注を誤設定した場合は文言の正誤に関わらず専用メッセージ
                string withEndnoteTokens = Regex.Replace(
                    malwareParagraphXml,
                    @"<w:endnoteReference\b[^>]*w:id=""([^""]+)""[^>]*>",
                    "[ENDNOTE_$1]");
                string endnoteCleanText = Regex.Replace(withEndnoteTokens, @"<[^>]+>", "");
                if (Regex.IsMatch(endnoteCleanText, @"マルウェア\[ENDNOTE_[^\]]+\]"))
                    return Miss("文末脚注が設定されています。");

                if (WordFindHelper.CountFootnoteDefinitionsInXml(xml) == 0)
                    return Miss("脚注が設定されていません。");

                var footnotesPartMatch = Regex.Match(xml, @"<pkg:part pkg:name=""/word/footnotes\.xml""[^>]*>.*?</pkg:part>", RegexOptions.Singleline);
                string footnotesXml = footnotesPartMatch.Value;
                var footnoteMatches = Regex.Matches(footnotesXml, @"<w:footnote\b[^>]*w:id=""([^""]+)""[^>]*>.*?</w:footnote>", RegexOptions.Singleline);
                var footnoteBodies = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (Match fn in footnoteMatches)
                {
                    string openTag = Regex.Match(fn.Value, @"<w:footnote\b[^>]*>").Value;
                    if (openTag.IndexOf("w:type=\"separator\"", StringComparison.Ordinal) >= 0
                        || openTag.IndexOf("w:type=\"continuationSeparator\"", StringComparison.Ordinal) >= 0
                        || openTag.IndexOf("w:type=\"continuationNotice\"", StringComparison.Ordinal) >= 0)
                        continue;

                    string bodyXmlFn = Regex.Replace(fn.Value, @"<w:footnoteRef\b[^>]*/>", "");
                    bodyXmlFn = Regex.Replace(bodyXmlFn, @"<w:footnoteRef\b[^>]*>.*?</w:footnoteRef>", "");
                    string bodyText = Regex.Replace(bodyXmlFn, @"<[^>]+>", "").Trim();
                    footnoteBodies[fn.Groups[1].Value] = bodyText;
                }

                string withFootnoteTokens = Regex.Replace(
                    malwareParagraphXml,
                    @"<w:footnoteReference\b[^>]*w:id=""([^""]+)""[^>]*>",
                    "[FOOTNOTE_$1]");
                string footnoteCleanText = Regex.Replace(withFootnoteTokens, @"<[^>]+>", "");
                Match attached = Regex.Match(footnoteCleanText, @"マルウェア\[FOOTNOTE_([^\]]+)\]");
                if (attached.Success)
                {
                    string footnoteId = attached.Groups[1].Value;
                    string actual = footnoteBodies.TryGetValue(footnoteId, out var body) ? body : string.Empty;
                    if (actual.Contains(expectedFootnote))
                        return true;
                    return Miss($"「{actual}」になっています。");
                }

                return Miss(missNotOnMalware);
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_03_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                var footnotePrMatch = Regex.Match(xml, @"<w:footnotePr\b[^>]*>.*?</w:footnotePr>", RegexOptions.Singleline);
                bool formatOk = footnotePrMatch.Success
                    && Regex.IsMatch(footnotePrMatch.Value, @"<w:numFmt\b[^>]*w:val=""decimalEnclosedCircle""");

                bool refCountOk = WordFindHelper.CountFootnoteReferencesInXml(xml) == 1;
                bool defCountOk = WordFindHelper.CountFootnoteDefinitionsInXml(xml) == 1;
                // 「適用」ではなく「挿入」した場合など、脚注本文にリテラル①が付くケースを落す
                bool markerOk = IsFootnoteNumberMarkerClean(xml);

                // 8-3は番号形式・件数・番号マーカーのみ。脚注文の正誤は8-2側で判定する。
                if (formatOk && refCountOk && defCountOk && markerOk)
                    return true;

                bool noted = false;
                if (!formatOk)
                {
                    WordScoreExplanation.Note("脚注の開始番号が「①」になっていません。");
                    noted = true;
                }
                if (!refCountOk || !defCountOk)
                {
                    WordScoreExplanation.Note("脚注の数が正しくありません。");
                    noted = true;
                }
                if (formatOk && refCountOk && defCountOk && !markerOk)
                {
                    WordScoreExplanation.Note("脚注に余分な番号が挿入されています。");
                    noted = true;
                }
                if (noted)
                    return false;
                return Miss("脚注の開始番号が「①」になっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>
        /// 8-3: 脚注定義の自動番号（w:footnoteRef）が1個だけで、
        /// 本文に丸数字リテラルが混入していないこと（「挿入」誤操作の①①等）。
        /// </summary>
        private static bool IsFootnoteNumberMarkerClean(string wordOpenXml)
        {
            var footnotesPartMatch = Regex.Match(
                wordOpenXml,
                @"<pkg:part pkg:name=""/word/footnotes\.xml""[^>]*>.*?</pkg:part>",
                RegexOptions.Singleline);
            if (!footnotesPartMatch.Success)
                return false;

            string footnotesXml = footnotesPartMatch.Value;
            var footnoteMatches = Regex.Matches(
                footnotesXml,
                @"<w:footnote\b[^>]*>.*?</w:footnote>",
                RegexOptions.Singleline);

            bool foundReal = false;
            foreach (Match fn in footnoteMatches)
            {
                string openTag = Regex.Match(fn.Value, @"<w:footnote\b[^>]*>").Value;
                if (openTag.IndexOf("w:type=\"separator\"", StringComparison.Ordinal) >= 0
                    || openTag.IndexOf("w:type=\"continuationSeparator\"", StringComparison.Ordinal) >= 0
                    || openTag.IndexOf("w:type=\"continuationNotice\"", StringComparison.Ordinal) >= 0)
                    continue;

                foundReal = true;
                if (Regex.Matches(fn.Value, @"<w:footnoteRef\b").Count != 1)
                    return false;

                if (Regex.IsMatch(
                    fn.Value,
                    @"<w:t\b[^>]*>[^<]*[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳\u2460-\u2473\u3251-\u325F]"))
                    return false;
            }

            return foundReal;
        }

        private bool CheckTask_1_8_04_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                var paragraphs = System.Text.RegularExpressions.Regex.Matches(xml, @"<w:p\b[^>]*>.*?</w:p>", System.Text.RegularExpressions.RegexOptions.Singleline);
                foreach (System.Text.RegularExpressions.Match p in paragraphs)
                {
                    string pXml = p.Value;
                    string pText = System.Text.RegularExpressions.Regex.Replace(pXml, @"<[^>]+>", "").Trim();
                    if (pText.Contains("参考文献"))
                    {
                        if (pXml.Contains("PAGEREF") || pXml.Contains("w:hyperlink")) continue;

                        bool containsSec = pText.Contains("§");
                        bool containsSym = pXml.Contains("<w:sym ");

                        if (containsSec || containsSym)
                            return true;
                        return Miss("見出し「参考文献」の先頭に「§」が挿入されていません。");
                    }
                }

                return Miss("見出し「参考文献」が見つかりません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_05_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                var partMatches = Regex.Matches(xml, @"<pkg:part pkg:name=""/word/diagrams/([^""]+)""[^>]*>.*?</pkg:part>", RegexOptions.Singleline);

                bool hasVennLayout = false;
                List<string> nodeTexts = new List<string>();
                bool hasRequiredTexts = false;

                foreach (Match match in partMatches)
                {
                    string partXml = match.Value;
                    string partName = match.Groups[1].Value;

                    if (partName.Contains("layout")
                        && partXml.Contains(@"urn:microsoft.com/office/officeart/2005/8/layout/venn1"))
                        hasVennLayout = true;
                    else if (partName.Contains("data"))
                    {
                        List<string> texts = GetSmartArtNodeTexts(partXml);
                        if (SmartArtNodeTextsMatchExpected(texts))
                        {
                            hasRequiredTexts = true;
                            nodeTexts = texts;
                        }
                        else if (!hasRequiredTexts
                            && (nodeTexts.Count == 0 || texts.Count == SmartArt805ExpectedTexts.Count))
                        {
                            // 失敗理由用: 3語ノードを優先して保持
                            nodeTexts = texts;
                        }
                    }
                }

                if (hasVennLayout && hasRequiredTexts)
                    return true;

                bool noted = false;
                if (!hasVennLayout)
                {
                    WordScoreExplanation.Note("Smart Art「基本ベン図」が挿入されていません。");
                    noted = true;
                }
                if (hasVennLayout && !hasRequiredTexts)
                {
                    WordScoreExplanation.Note(BuildSmartArtTextFailReason(nodeTexts));
                    noted = true;
                }
                if (noted)
                    return false;
                return Miss("Smart Art「基本ベン図」が指定どおりになっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private static readonly HashSet<string> SmartArt805ExpectedTexts = new HashSet<string>(StringComparer.Ordinal)
        {
            "機密性", "完全性", "可用性"
        };

        private static List<string> GetSmartArtNodeTexts(string dataXml)
        {
            var texts = new List<string>();
            if (string.IsNullOrEmpty(dataXml))
                return texts;

            // 基本ベン図の文言ノードは type 省略（OOXML 既定=node）。
            // type="doc" / parTrans / sibTrans / pres は除外する。
            foreach (Match pt in Regex.Matches(
                dataXml,
                @"<dgm:pt\b([^>]*)>(.*?)</dgm:pt>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                string attrs = pt.Groups[1].Value;
                Match typeMatch = Regex.Match(attrs, @"\btype\s*=\s*""([^""]+)""", RegexOptions.IgnoreCase);
                if (typeMatch.Success)
                {
                    string type = typeMatch.Groups[1].Value;
                    if (!string.Equals(type, "node", StringComparison.OrdinalIgnoreCase))
                        continue;
                }

                var sb = new System.Text.StringBuilder();
                foreach (Match t in Regex.Matches(
                    pt.Groups[2].Value,
                    @"<a:t\b[^>]*>(.*?)</a:t>",
                    RegexOptions.Singleline))
                {
                    sb.Append(t.Groups[1].Value);
                }
                texts.Add(sb.ToString().Trim());
            }

            return texts;
        }

        private static bool SmartArtNodeTextsMatchExpected(List<string> nodeTexts)
        {
            if (nodeTexts == null || nodeTexts.Count != SmartArt805ExpectedTexts.Count)
                return false;
            var actual = new HashSet<string>(nodeTexts, StringComparer.Ordinal);
            return actual.SetEquals(SmartArt805ExpectedTexts);
        }

        private static string BuildSmartArtTextFailReason(List<string> nodeTexts)
        {
            if (nodeTexts == null || nodeTexts.Count == 0)
                return "すべて入力されていません。";

            bool allEmpty = true;
            for (int i = 0; i < nodeTexts.Count; i++)
            {
                if (!string.IsNullOrEmpty(nodeTexts[i]))
                {
                    allEmpty = false;
                    break;
                }
            }
            if (allEmpty)
                return "すべて入力されていません。";

            var parts = new List<string>(nodeTexts.Count);
            for (int i = 0; i < nodeTexts.Count; i++)
            {
                string display = string.IsNullOrEmpty(nodeTexts[i]) ? "（空）" : nodeTexts[i];
                parts.Add($"「{display}」");
            }
            return string.Join(string.Empty, parts) + "と入力されています。";
        }

        private bool CheckTask_1_8_06_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                var partMatches = System.Text.RegularExpressions.Regex.Matches(xml, @"<pkg:part pkg:name=""/word/diagrams/([^""]+)""[^>]*>.*?</pkg:part>", System.Text.RegularExpressions.RegexOptions.Singleline);

                foreach (System.Text.RegularExpressions.Match match in partMatches)
                {
                    string partXml = match.Value;
                    string partName = match.Groups[1].Value;

                    if (partName.Contains("colors")
                        && partXml.Contains(@"urn:microsoft.com/office/officeart/2005/8/colors/colorful2"))
                        return true;
                }

                return Miss("Smart Artの色が「カラフル」の「アクセント2から3」になっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_07_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                var alternates = Regex.Matches(xml, @"<mc:AlternateContent\b[^>]*>.*?</mc:AlternateContent>", RegexOptions.Singleline);
                bool foundTopShape = false;
                bool foundTextOnlyLink = false;

                foreach (Match alt in alternates)
                {
                    string altXml = alt.Value;
                    if (!Regex.IsMatch(altXml, @"<w:t\b[^>]*>TOP</w:t>"))
                        continue;

                    foundTopShape = true;
                    if (HasShapeTopHyperlink(altXml, xml))
                        return true;
                    if (HasTextTopHyperlink(altXml))
                        foundTextOnlyLink = true;
                }

                if (foundTextOnlyLink)
                    return Miss("図形内の文字にリンクが設定されています。");
                if (!foundTopShape)
                    return Miss("「TOP」の図形が見つかりません。");
                return Miss("「TOP」の図形に文頭へのリンクが設定されていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>
        /// 図形全体への文頭リンク（VML href または DrawingML hlinkClick → #_top）。
        /// </summary>
        private static bool HasShapeTopHyperlink(string shapeXml, string wordOpenXml)
        {
            if (Regex.IsMatch(shapeXml, @"<v:shape\b[^>]*\bhref\s*=\s*""#_top""", RegexOptions.IgnoreCase | RegexOptions.Singleline))
                return true;

            Match click = Regex.Match(
                shapeXml,
                @"<a:hlinkClick\b[^>]*\br:id\s*=\s*""(?<id>[^""]+)""",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!click.Success)
                return false;

            string rid = click.Groups["id"].Value;
            return Regex.IsMatch(
                       wordOpenXml,
                       @"<Relationship\b[^>]*\bId\s*=\s*""" + Regex.Escape(rid) + @"""[^>]*\bTarget\s*=\s*""#_top""",
                       RegexOptions.IgnoreCase | RegexOptions.Singleline)
                   || Regex.IsMatch(
                       wordOpenXml,
                       @"<Relationship\b[^>]*\bTarget\s*=\s*""#_top""[^>]*\bId\s*=\s*""" + Regex.Escape(rid) + @"""",
                       RegexOptions.IgnoreCase | RegexOptions.Singleline);
        }

        /// <summary>
        /// 図形内テキスト「TOP」へのハイパーリンク（よくある誤操作）。
        /// </summary>
        private static bool HasTextTopHyperlink(string shapeXml)
        {
            foreach (Match txbx in Regex.Matches(
                shapeXml,
                @"<w:txbxContent\b[^>]*>.*?</w:txbxContent>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                string box = txbx.Value;
                if (!Regex.IsMatch(box, @"<w:t\b[^>]*>TOP</w:t>"))
                    continue;

                if (Regex.IsMatch(
                    box,
                    @"<w:hyperlink\b[^>]*\bw:anchor\s*=\s*""_top""",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline))
                    return true;
            }

            return false;
        }

        private static bool Miss(string reason)
        {
            WordScoreExplanation.Note(reason);
            return false;
        }

        private string GetCurrentWordFilePath()
        {
            Application wordApp = null;
            try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); if (wordApp.ActiveDocument != null) return wordApp.ActiveDocument.FullName; return null; }
            catch (COMException) { return null; }
            finally { if (wordApp != null) Marshal.ReleaseComObject(wordApp); }
        }
    }
}
