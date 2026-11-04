using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Libraries;
using Microsoft.Office.Interop.Word;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題第1回 プロジェクト6。6-1〜6-4 は最終状態のみ（ログ・破壊ゲートなし）。
    /// </summary>
    public class WordChecker1_6_PV1
    {
        private const string MailAddress = "star@example.jp";
        private const string MailtoAddress = "mailto:star@example.jp";
        private const string CommentParentPhrase = "上映時間を直してください";
        private const string CommentReplyPhrase = "夜の部は19時開始です";
        private const string ShapeTextAnchor = "秋の夜空を、";
        private const string ShapeTextInsert = "ぜひ";
        private const double Model3dXRotationTargetDegrees = 25d;
        private const double Model3dXRotationToleranceDegrees = 2d;

        public bool CheckTask_1_6_01() => RunCheck(CheckTask_1_6_01_Impl);
        public bool CheckTask_1_6_02() => RunCheck(CheckTask_1_6_02_Impl);
        public bool CheckTask_1_6_03() => RunCheck(CheckTask_1_6_03_Impl);
        public bool CheckTask_1_6_04() => RunCheck(CheckTask_1_6_04_Impl);

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

        private bool CheckTask_1_6_01_Impl(string filePath)
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

                if (HyperlinkMailtoMatches(document, xml))
                    return true;

                return Miss("「star@example.jp」に正しい電子メールリンクが設定されていません。");
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool HyperlinkMailtoMatches(Document document, string xml)
        {
            if (DocumentHyperlinksMailtoOk(document))
                return true;

            return XmlMailtoHyperlinkOk(xml);
        }

        private static bool DocumentHyperlinksMailtoOk(Document document)
        {
            Hyperlinks links = null;
            try
            {
                links = document.Hyperlinks;
                int count = links.Count;
                for (int i = 1; i <= count; i++)
                {
                    Hyperlink link = null;
                    try
                    {
                        link = links[i];
                        if (!IsMailtoStarLink(link))
                            continue;

                        string display = string.Empty;
                        try { display = link.TextToDisplay ?? string.Empty; } catch { }
                        if (display.IndexOf("mailto:", StringComparison.OrdinalIgnoreCase) >= 0)
                            continue;

                        if (display.IndexOf(MailAddress, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;

                        Range range = null;
                        try
                        {
                            range = link.Range;
                            string text = (range?.Text ?? string.Empty).Trim('\r', '\n', '\a', ' ');
                            if (text.IndexOf("mailto:", StringComparison.OrdinalIgnoreCase) >= 0)
                                continue;
                            if (text.IndexOf(MailAddress, StringComparison.OrdinalIgnoreCase) >= 0)
                                return true;
                        }
                        finally
                        {
                            if (range != null)
                                Marshal.ReleaseComObject(range);
                        }
                    }
                    catch { }
                    finally
                    {
                        if (link != null)
                            Marshal.ReleaseComObject(link);
                    }
                }
            }
            catch { }
            finally
            {
                if (links != null)
                    Marshal.ReleaseComObject(links);
            }

            return false;
        }

        private static bool IsMailtoStarLink(Hyperlink link)
        {
            string address = string.Empty;
            try { address = link.Address ?? string.Empty; } catch { }
            address = address.Trim();
            if (address.StartsWith(MailtoAddress, StringComparison.OrdinalIgnoreCase))
                return true;

            return string.Equals(address, MailAddress, StringComparison.OrdinalIgnoreCase)
                && address.IndexOf('@') >= 0;
        }

        private static bool XmlMailtoHyperlinkOk(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            if (xml.IndexOf(MailtoAddress, StringComparison.OrdinalIgnoreCase) < 0
                && xml.IndexOf("mailto:" + MailAddress, StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }

            if (xml.IndexOf(MailAddress, StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            return xml.IndexOf("mailto:" + MailAddress, StringComparison.OrdinalIgnoreCase) >= 0
                || Regex.IsMatch(xml, @"Target=""mailto:" + Regex.Escape(MailAddress) + @"""", RegexOptions.IgnoreCase);
        }

        private bool CheckTask_1_6_02_Impl(string filePath)
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

                if (DocumentHasCommentReply(document, xml, CommentParentPhrase, CommentReplyPhrase))
                    return true;

                return Miss("「上映時間を直してください」への指定の返信がありません。");
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool DocumentHasCommentReply(
            Document document,
            string xml,
            string parentPhrase,
            string replyPhrase)
        {
            if (TryFindCommentReplyCom(document, parentPhrase, replyPhrase))
                return true;

            return XmlContainsCommentReply(xml, parentPhrase, replyPhrase);
        }

        private static bool TryFindCommentReplyCom(Document document, string parentPhrase, string replyPhrase)
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
                        if (!CommentRelatesToParent(comment, parentPhrase))
                            continue;

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
                                    string text = string.Empty;
                                    try { text = reply.Range?.Text ?? string.Empty; } catch { }
                                    if (NormalizeCommentBody(text).Contains(replyPhrase))
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

        private static bool CommentRelatesToParent(Comment comment, string parentPhrase)
        {
            string body = string.Empty;
            try { body = comment.Range?.Text ?? string.Empty; } catch { }
            if (NormalizeCommentBody(body).Contains(parentPhrase))
                return true;

            Range scope = null;
            try
            {
                scope = comment.Scope;
                string scopeText = scope?.Text ?? string.Empty;
                return NormalizeCommentBody(scopeText).Contains(parentPhrase);
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

        private static bool XmlContainsCommentReply(string xml, string parentPhrase, string replyPhrase)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            if (xml.IndexOf(replyPhrase, StringComparison.Ordinal) < 0)
                return false;

            return xml.IndexOf(parentPhrase, StringComparison.Ordinal) >= 0;
        }

        private static string NormalizeCommentBody(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;
            return text.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\a", string.Empty).Trim();
        }

        private bool CheckTask_1_6_03_Impl(string filePath)
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

                if (Model3dXRotationIsAbout25(xml) || Model3dXRotationIsAbout25Com(document))
                    return true;

                return Miss("記念グラスの3D画像のX方向の回転が25°になっていません。");
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool Model3dXRotationIsAbout25(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            const int unitsPerDegree = 60000;
            long targetUnits = (long)Math.Round(Model3dXRotationTargetDegrees * unitsPerDegree);
            long toleranceUnits = (long)Math.Round(Model3dXRotationToleranceDegrees * unitsPerDegree);

            // Word 3D: am3d:trans / am3d:rot ax="1500000" (= 25° × 60000)
            foreach (Match match in Regex.Matches(
                xml,
                @"am3d:rot\b[^>]*\bax\s*=\s*[""'](-?\d+)[""']",
                RegexOptions.IgnoreCase))
            {
                if (Am3dAxUnitsNearTarget(match.Groups[1].Value, targetUnits, toleranceUnits))
                    return true;
            }

            // 名前空間プレフィックスが落ちる場合のフォールバック（model3d ブロック内の rot ax）
            foreach (Match block in Regex.Matches(
                xml,
                @"am3d:model3d\b[\s\S]*?(?:</am3d:model3d>|/>)",
                RegexOptions.IgnoreCase))
            {
                foreach (Match rot in Regex.Matches(
                    block.Value,
                    @"\brot\b[^>]*\bax\s*=\s*[""'](-?\d+)[""']",
                    RegexOptions.IgnoreCase))
                {
                    if (Am3dAxUnitsNearTarget(rot.Groups[1].Value, targetUnits, toleranceUnits))
                        return true;
                }
            }

            return false;
        }

        private static bool Am3dAxUnitsNearTarget(string rawAx, long targetUnits, long toleranceUnits)
        {
            if (!long.TryParse(rawAx, out long axUnits))
                return false;

            return Math.Abs(axUnits - targetUnits) <= toleranceUnits;
        }

        private static bool Model3dXRotationIsAbout25Com(Document document)
        {
            InlineShapes inlines = null;
            try
            {
                inlines = document.InlineShapes;
                int count = inlines.Count;
                for (int i = 1; i <= count; i++)
                {
                    InlineShape inline = null;
                    try
                    {
                        inline = inlines[i];
                        try
                        {
                            dynamic dyn = inline;
                            dynamic model3d = dyn.Model3D;
                            if (model3d == null)
                                continue;

                            float rotX = (float)model3d.RotationX;
                            if (Math.Abs(rotX - (float)Model3dXRotationTargetDegrees) <= (float)Model3dXRotationToleranceDegrees)
                                return true;
                        }
                        catch { }
                    }
                    finally
                    {
                        if (inline != null)
                            Marshal.ReleaseComObject(inline);
                    }
                }
            }
            catch { }
            finally
            {
                if (inlines != null)
                    Marshal.ReleaseComObject(inlines);
            }

            return false;
        }

        private bool CheckTask_1_6_04_Impl(string filePath)
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

                if (ShapeContainsInsertedText(document))
                    return true;

                string xml;
                try { xml = document.WordOpenXML ?? string.Empty; }
                catch { xml = string.Empty; }

                if (ShapeXmlContainsInsertedText(xml))
                    return true;

                return Miss("図形内の「秋の夜空を、」の後に「ぜひ」が追加されていません。");
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool ShapeContainsInsertedText(Document document)
        {
            foreach (Microsoft.Office.Interop.Word.Shape shape in document.Shapes)
            {
                try
                {
                    if (ShapeTextHasInsertAfterAnchor(shape))
                        return true;
                }
                catch { }
                finally
                {
                    Marshal.ReleaseComObject(shape);
                }
            }

            return false;
        }

        private static bool ShapeTextHasInsertAfterAnchor(Microsoft.Office.Interop.Word.Shape shape)
        {
            TextFrame frame = null;
            Range range = null;
            try
            {
                frame = shape.TextFrame;
                if (frame == null || frame.HasText == 0)
                    return false;

                range = frame.TextRange;
                string text = NormalizeShapeText(range?.Text ?? string.Empty);
                return TextHasInsertImmediatelyAfterAnchor(text);
            }
            finally
            {
                if (range != null)
                    Marshal.ReleaseComObject(range);
                if (frame != null)
                    Marshal.ReleaseComObject(frame);
            }
        }

        private static bool ShapeXmlContainsInsertedText(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            int idx = xml.IndexOf(ShapeTextAnchor, StringComparison.Ordinal);
            if (idx < 0)
                return false;

            int start = Math.Max(0, idx - 500);
            int end = Math.Min(xml.Length, idx + 500);
            string window = xml.Substring(start, end - start);
            string plain = Regex.Replace(window, @"<[^>]+>", string.Empty);
            plain = NormalizeShapeText(plain);
            return TextHasInsertImmediatelyAfterAnchor(plain);
        }

        private static bool TextHasInsertImmediatelyAfterAnchor(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            int idx = text.IndexOf(ShapeTextAnchor, StringComparison.Ordinal);
            if (idx < 0)
                return false;

            string after = text.Substring(idx + ShapeTextAnchor.Length).TrimStart(' ', '　', '\t');
            return after.StartsWith(ShapeTextInsert, StringComparison.Ordinal);
        }

        private static string NormalizeShapeText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;
            return text.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\a", string.Empty);
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
                if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                    doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
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