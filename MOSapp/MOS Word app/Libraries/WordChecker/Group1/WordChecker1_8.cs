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

                var documentPartMatch = Regex.Match(xml, @"<pkg:part pkg:name=""/word/document\.xml""[^>]*>.*?</pkg:part>", RegexOptions.Singleline);
                string bodyXml = documentPartMatch.Success ? documentPartMatch.Value : xml;
                var paragraphs = Regex.Matches(bodyXml, @"<w:p\b[^>]*>.*?</w:p>", RegexOptions.Singleline);
                bool foundMalwareParagraph = false;

                foreach (Match p in paragraphs)
                {
                    string pXml = p.Value;
                    string pTextOnly = Regex.Replace(pXml, @"<[^>]+>", "");
                    if (!pTextOnly.Contains("マルウェア"))
                        continue;

                    foundMalwareParagraph = true;
                    string replacedXml = Regex.Replace(pXml, @"<w:footnoteReference\b[^>]*w:id=""([^""]+)""[^>]*>", "[FOOTNOTE_$1]");
                    string cleanText = Regex.Replace(replacedXml, @"<[^>]+>", "");

                    Match attached = Regex.Match(cleanText, @"マルウェア\[FOOTNOTE_([^\]]+)\]");
                    if (attached.Success)
                    {
                        string footnoteId = attached.Groups[1].Value;
                        string actual = footnoteBodies.TryGetValue(footnoteId, out var body) ? body : string.Empty;
                        if (actual.Contains(expectedFootnote))
                            return true;
                        return Miss($"脚注が「{actual}」になっています。");
                    }

                    return Miss(missNotOnMalware);
                }

                if (!foundMalwareParagraph)
                    return Miss("「マルウェア」が見つかりません。");
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

                var footnotePrMatch = System.Text.RegularExpressions.Regex.Match(xml, @"<w:footnotePr\b[^>]*>.*?</w:footnotePr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                bool formatOk = footnotePrMatch.Success
                    && System.Text.RegularExpressions.Regex.IsMatch(footnotePrMatch.Value, @"<w:numFmt\b[^>]*w:val=""decimalEnclosedCircle""");

                int refCount = WordFindHelper.CountFootnoteReferencesInXml(xml);
                bool refCountOk = refCount == 1;
                bool defCountOk = WordFindHelper.CountFootnoteDefinitionsInXml(xml) == 1;
                bool bodyOk = IsFootnoteBodyValidForTask803(xml);

                if (formatOk && refCountOk && defCountOk && bodyOk)
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
                if (formatOk && refCountOk && defCountOk && !bodyOk)
                {
                    WordScoreExplanation.Note("脚注の内容が正しくありません。");
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
        /// 8-3: 脚注定義本文に w:footnoteRef が 1 個のみで、丸数字リテラルや余分な記号がなく、
        /// 期待テキストだけが残ることを検証する。
        /// </summary>
        private static bool IsFootnoteBodyValidForTask803(string wordOpenXml)
        {
            var footnotesPartMatch = System.Text.RegularExpressions.Regex.Match(
                wordOpenXml,
                @"<pkg:part pkg:name=""/word/footnotes\.xml""[^>]*>.*?</pkg:part>",
                System.Text.RegularExpressions.RegexOptions.Singleline);
            if (!footnotesPartMatch.Success)
                return false;

            string footnotesXml = footnotesPartMatch.Value;
            var footnoteMatches = System.Text.RegularExpressions.Regex.Matches(
                footnotesXml,
                @"<w:footnote\b[^>]*w:id=""([^""]+)""[^>]*>.*?</w:footnote>",
                System.Text.RegularExpressions.RegexOptions.Singleline);

            string targetFootnoteXml = null;
            foreach (System.Text.RegularExpressions.Match fn in footnoteMatches)
            {
                string openTag = System.Text.RegularExpressions.Regex.Match(fn.Value, @"<w:footnote\b[^>]*>").Value;
                if (openTag.IndexOf("w:type=\"separator\"", StringComparison.Ordinal) >= 0
                    || openTag.IndexOf("w:type=\"continuationSeparator\"", StringComparison.Ordinal) >= 0
                    || openTag.IndexOf("w:type=\"continuationNotice\"", StringComparison.Ordinal) >= 0)
                    continue;

                string fnText = System.Text.RegularExpressions.Regex.Replace(fn.Value, @"<[^>]+>", "");
                if (fnText.Contains(FootnoteTargetText803))
                {
                    targetFootnoteXml = fn.Value;
                    break;
                }
            }

            if (string.IsNullOrEmpty(targetFootnoteXml))
                return false;

            if (System.Text.RegularExpressions.Regex.Matches(targetFootnoteXml, @"<w:footnoteRef\b").Count != 1)
                return false;

            if (System.Text.RegularExpressions.Regex.IsMatch(
                targetFootnoteXml,
                @"<w:t\b[^>]*>[^<]*[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳\u2460-\u2473\u3251-\u325F]"))
                return false;

            string bodyText = System.Text.RegularExpressions.Regex.Replace(
                targetFootnoteXml, @"<w:footnoteRef\b[^>]*/>", "");
            bodyText = System.Text.RegularExpressions.Regex.Replace(
                bodyText, @"<w:footnoteRef\b[^>]*>.*?</w:footnoteRef>", "");
            bodyText = System.Text.RegularExpressions.Regex.Replace(bodyText, @"<[^>]+>", "");
            return string.Equals(bodyText.Trim(), FootnoteTargetText803, StringComparison.Ordinal);
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

                var partMatches = System.Text.RegularExpressions.Regex.Matches(xml, @"<pkg:part pkg:name=""/word/diagrams/([^""]+)""[^>]*>.*?</pkg:part>", System.Text.RegularExpressions.RegexOptions.Singleline);

                bool hasVennLayout = false;
                bool hasRequiredTexts = false;

                foreach (System.Text.RegularExpressions.Match match in partMatches)
                {
                    string partXml = match.Value;
                    string partName = match.Groups[1].Value;

                    if (partName.Contains("layout"))
                    {
                        if (partXml.Contains(@"urn:microsoft.com/office/officeart/2005/8/layout/venn1"))
                            hasVennLayout = true;
                    }
                    else if (partName.Contains("data"))
                    {
                        if (partXml.Contains("機密性") && partXml.Contains("完全性") && partXml.Contains("可用性"))
                            hasRequiredTexts = true;
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
                    WordScoreExplanation.Note("「機密性」「完全性」「可用性」が入力されていません。");
                    noted = true;
                }
                if (noted)
                    return false;
                return Miss("Smart Art「基本ベン図」が指定どおりになっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
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

                var alternates = System.Text.RegularExpressions.Regex.Matches(xml, @"<mc:AlternateContent\b[^>]*>.*?</mc:AlternateContent>", System.Text.RegularExpressions.RegexOptions.Singleline);

                foreach (System.Text.RegularExpressions.Match alt in alternates)
                {
                    string altXml = alt.Value;
                    if (altXml.Contains("TOP") && altXml.Contains("href=\"#_top\""))
                        return true;
                }

                return Miss("「TOP」の図形に文頭へのリンクが設定されていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
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
