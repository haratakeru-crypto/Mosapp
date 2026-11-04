using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Libraries;
using Microsoft.Office.Interop.Word;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題第1回 プロジェクト8。8-1〜8-4 は最終状態のみ。8-5 は教材1-1と同じ編集記号ログ。
    /// </summary>
    public class WordChecker1_8_PV1
    {
        private const int VariantSetNo = 1;
        private const string IconAnchorSentence = "食器は子どもが自分で選べるように、低い棚に置きます。";
        private const string CommentDeletePhrase = "メニューを増やす";
        private const string MenuHeadingPhrase = "今週の献立";
        private const long IconExtentEmu = 914400L;
        private const long IconExtentToleranceEmu = 80000L;
        private const int MenuSectionMinParagraphCount = 8;

        public bool CheckTask_1_8_01() => RunCheck(CheckTask_1_8_01_Impl);
        public bool CheckTask_1_8_02() => RunCheck(CheckTask_1_8_02_Impl);
        public bool CheckTask_1_8_03() => RunCheck(CheckTask_1_8_03_Impl);
        public bool CheckTask_1_8_04() => RunCheck(CheckTask_1_8_04_Impl);
        public bool CheckTask_1_8_05() => RunCheck(CheckTask_1_8_05_Impl);

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
            Application wordApp = null;
            Document document = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                string xml;
                try { xml = document.WordOpenXML ?? string.Empty; }
                catch { xml = string.Empty; }

                if (IconInsertedAfterAnchor(xml))
                    return true;

                if (IconInsertedAfterAnchorCom(document))
                    return true;

                return Miss("指定位置にアイコンが挿入されていません（サイズ変更なし）。");
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_8_02_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                string xml;
                try { xml = document.WordOpenXML ?? string.Empty; }
                catch { xml = string.Empty; }

                if (FootnoteNumberFormatIsEnclosedCircle(xml))
                    return true;

                return Miss("脚注の番号書式が［①, ②, ③, …］になっていません。");
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_8_03_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                if (DocumentHasCommentPhrase(document))
                    return Miss("「メニューを増やす」のコメントが削除されていません。");

                string xml;
                try { xml = document.WordOpenXML ?? string.Empty; }
                catch { xml = string.Empty; }

                if (XmlHasCommentPhrase(xml))
                    return Miss("「メニューを増やす」のコメントが削除されていません。");

                return true;
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_8_04_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                string xml;
                try { xml = document.WordOpenXML ?? string.Empty; }
                catch { xml = string.Empty; }

                if (MenuTableConvertedToParagraphs(xml))
                    return true;

                return Miss("「今週の献立」の下の表が段落記号区切りで解除されていません。");
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_8_05_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                // 教材1-1と同じ: トグル証跡2回以上 かつ FinalOn/FinalOff の last-wins が FinalOn。
                // View.ShowAll は再オープンで消えるため採点式には使わない。類題ログは教材と混ざらない。
                bool showAllExecutedTwice = LogReader.HasVariantTaskEvidenceAtLeast(
                    VariantSetNo, 8, 5, "ShowAll", 2);
                if (!showAllExecutedTwice)
                    return Miss("編集記号の表示について、必要な操作の記録がありません。");

                if (LogReader.IsLatestVariantShowAllFinalOn(VariantSetNo, 8, 5))
                    return true;

                return Miss("編集記号が表示されていません。");
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool IconInsertedAfterAnchor(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            int anchorIndex = xml.IndexOf(IconAnchorSentence, StringComparison.Ordinal);
            if (anchorIndex < 0)
                return false;

            int windowLength = Math.Min(12000, xml.Length - anchorIndex);
            if (windowLength <= 0)
                return false;

            string window = xml.Substring(anchorIndex, windowLength);
            if (window.IndexOf("<w:drawing", StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            bool looksLikeIcon = window.IndexOf("drawing/2016/SVG", StringComparison.OrdinalIgnoreCase) >= 0
                || window.IndexOf(":SVG", StringComparison.OrdinalIgnoreCase) >= 0;

            if (!looksLikeIcon)
                return false;

            return IconExtentMatchesDefault(window);
        }

        private static bool IconExtentMatchesDefault(string window)
        {
            foreach (Match match in Regex.Matches(window, @"<wp:extent\b[^>]*cx=""(?<cx>\d+)""[^>]*cy=""(?<cy>\d+)""", RegexOptions.IgnoreCase))
            {
                if (TryExtentMatchesDefault(match.Groups["cx"].Value, match.Groups["cy"].Value))
                    return true;
            }

            foreach (Match match in Regex.Matches(window, @"<wp:extent\b[^>]*cy=""(?<cy>\d+)""[^>]*cx=""(?<cx>\d+)""", RegexOptions.IgnoreCase))
            {
                if (TryExtentMatchesDefault(match.Groups["cx"].Value, match.Groups["cy"].Value))
                    return true;
            }

            return false;
        }

        private static bool TryExtentMatchesDefault(string cxRaw, string cyRaw)
        {
            if (!long.TryParse(cxRaw, out long cx) || !long.TryParse(cyRaw, out long cy))
                return false;

            return Math.Abs(cx - IconExtentEmu) <= IconExtentToleranceEmu
                && Math.Abs(cy - IconExtentEmu) <= IconExtentToleranceEmu;
        }

        private static bool IconInsertedAfterAnchorCom(Document document)
        {
            Paragraph anchorPara = FindParagraphContaining(document, IconAnchorSentence);
            if (anchorPara == null)
                return false;

            int afterStart = anchorPara.Range.End;
            Marshal.ReleaseComObject(anchorPara);

            foreach (Paragraph paragraph in document.Paragraphs)
            {
                try
                {
                    if (paragraph.Range.Start < afterStart)
                        continue;

                    if (paragraph.Range.Start > afterStart + 800)
                        break;

                    InlineShapes shapes = null;
                    try
                    {
                        shapes = paragraph.Range.InlineShapes;
                        int count = shapes.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            InlineShape shape = null;
                            try
                            {
                                shape = shapes[i];
                                float w = 0f;
                                float h = 0f;
                                try { w = (float)shape.Width; } catch { }
                                try { h = (float)shape.Height; } catch { }
                                if (w <= 0f || h <= 0f)
                                    continue;

                                if (Math.Abs(w - h) <= 2f && w >= 50f && w <= 90f)
                                    return true;
                            }
                            finally
                            {
                                if (shape != null)
                                    Marshal.ReleaseComObject(shape);
                            }
                        }
                    }
                    finally
                    {
                        if (shapes != null)
                            Marshal.ReleaseComObject(shapes);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(paragraph);
                }
            }

            return false;
        }

        private static bool FootnoteNumberFormatIsEnclosedCircle(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            if (Regex.IsMatch(
                xml,
                @"<w:numFmt\b[^>]*w:val=""decimalEnclosedCircle""",
                RegexOptions.IgnoreCase))
            {
                return true;
            }

            Match footnotePr = Regex.Match(
                xml,
                @"<w:footnotePr\b[^>]*>.*?</w:footnotePr>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (!footnotePr.Success)
                return false;

            return Regex.IsMatch(
                footnotePr.Value,
                @"<w:numFmt\b[^>]*w:val=""decimalEnclosedCircle""",
                RegexOptions.IgnoreCase);
        }

        private static bool DocumentHasCommentPhrase(Document document)
        {
            Comments comments = null;
            try
            {
                comments = document.Comments;
                int count = comments.Count;
                for (int i = 1; i <= count; i++)
                {
                    Comment comment = null;
                    try
                    {
                        comment = comments[i];
                        if (CommentContainsPhrase(comment, CommentDeletePhrase))
                            return true;

                        Comments replies = null;
                        try
                        {
                            replies = comment.Replies;
                            if (replies == null)
                                continue;

                            int replyCount = replies.Count;
                            for (int r = 1; r <= replyCount; r++)
                            {
                                Comment reply = null;
                                try
                                {
                                    reply = replies[r];
                                    if (CommentContainsPhrase(reply, CommentDeletePhrase))
                                        return true;
                                }
                                finally
                                {
                                    if (reply != null)
                                        Marshal.ReleaseComObject(reply);
                                }
                            }
                        }
                        finally
                        {
                            if (replies != null)
                                Marshal.ReleaseComObject(replies);
                        }
                    }
                    finally
                    {
                        if (comment != null)
                            Marshal.ReleaseComObject(comment);
                    }
                }
            }
            catch { }
            finally
            {
                if (comments != null)
                    Marshal.ReleaseComObject(comments);
            }

            return false;
        }

        private static bool CommentContainsPhrase(Comment comment, string phrase)
        {
            string text = string.Empty;
            try { text = comment.Range?.Text ?? string.Empty; } catch { }
            text = NormalizeCommentText(text);
            if (text.IndexOf(phrase, StringComparison.Ordinal) >= 0)
                return true;

            Range scope = null;
            try
            {
                scope = comment.Scope;
                string scopeText = scope?.Text ?? string.Empty;
                return NormalizeCommentText(scopeText).IndexOf(phrase, StringComparison.Ordinal) >= 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (scope != null)
                    Marshal.ReleaseComObject(scope);
            }
        }

        private static bool XmlHasCommentPhrase(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            foreach (Match part in Regex.Matches(
                xml,
                @"<pkg:part pkg:name=""/word/comments[^""]*""[^>]*>.*?</pkg:part>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                if (part.Value.IndexOf(CommentDeletePhrase, StringComparison.Ordinal) >= 0)
                    return true;
            }

            return false;
        }

        private static string NormalizeCommentText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;
            return text.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\a", string.Empty).Trim();
        }

        private static bool MenuTableConvertedToParagraphs(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            string bodyXml = ExtractDocumentBodyXml(xml);
            int headingIndex = bodyXml.IndexOf(MenuHeadingPhrase, StringComparison.Ordinal);
            if (headingIndex < 0)
                return false;

            int sliceLength = Math.Min(80000, bodyXml.Length - headingIndex);
            if (sliceLength <= 0)
                return false;

            string section = bodyXml.Substring(headingIndex, sliceLength);
            if (section.IndexOf("<w:tbl", StringComparison.OrdinalIgnoreCase) >= 0)
                return false;

            string plain = Regex.Replace(section, @"<[^>]+>", string.Empty);
            if (plain.IndexOf("月曜", StringComparison.Ordinal) < 0
                || plain.IndexOf("みそ汁", StringComparison.Ordinal) < 0)
            {
                return false;
            }

            int paragraphCount = Regex.Matches(section, @"<w:p\b", RegexOptions.IgnoreCase).Count;
            return paragraphCount >= MenuSectionMinParagraphCount;
        }

        private static string ExtractDocumentBodyXml(string wordOpenXml)
        {
            Match match = Regex.Match(
                wordOpenXml,
                @"<pkg:part pkg:name=""/word/document.xml""[^>]*>.*?</pkg:part>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            return match.Success ? match.Value : wordOpenXml;
        }

        private static Paragraph FindParagraphContaining(Document document, string part)
        {
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                if (ParagraphContainsText(paragraph, part))
                    return paragraph;
                Marshal.ReleaseComObject(paragraph);
            }

            return null;
        }

        private static bool ParagraphContainsText(Paragraph paragraph, string part)
        {
            try
            {
                return (paragraph.Range.Text ?? string.Empty).IndexOf(part, StringComparison.Ordinal) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private static Application GetWordApp()
        {
            try { return (Application)Marshal.GetActiveObject("Word.Application"); }
            catch { return null; }
        }

        private static Document GetDocument(Application wordApp, string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            foreach (Document doc in wordApp.Documents)
            {
                if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                    || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                {
                    return doc;
                }
            }

            return null;
        }

        private static string GetCurrentWordFilePath()
        {
            Application wordApp = null;
            try
            {
                wordApp = GetWordApp();
                return wordApp?.ActiveDocument?.FullName;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                if (wordApp != null)
                    Marshal.ReleaseComObject(wordApp);
            }
        }

        private static bool Miss(string reason)
        {
            WordScoreExplanation.Note(reason);
            return false;
        }
    }
}
