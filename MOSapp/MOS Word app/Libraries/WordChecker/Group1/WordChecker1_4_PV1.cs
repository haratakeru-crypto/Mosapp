using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Libraries;
using Microsoft.Office.Interop.Word;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題第1回 プロジェクト4。4-1〜4-4 は最終状態のみ（ログ・破壊ゲートなし）。
    /// </summary>
    public class WordChecker1_4_PV1
    {
        private const float LeftIndentCharTarget = 3f;
        private const float LeftIndentCharTolerance = 0.15f;

        private static readonly string[] FlowParagraphPrefixes =
        {
            "開場は",
            "演奏は",
            "終演後は"
        };

        public bool CheckTask_1_4_01() => RunCheck(CheckTask_1_4_01_Impl);
        public bool CheckTask_1_4_02() => RunCheck(CheckTask_1_4_02_Impl);
        public bool CheckTask_1_4_03() => RunCheck(CheckTask_1_4_03_Impl);
        public bool CheckTask_1_4_04() => RunCheck(CheckTask_1_4_04_Impl);

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

        private bool CheckTask_1_4_01_Impl(string filePath)
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

                if (document.Sections.Count < 1)
                    return Miss("ページ罫線を確認できません。");

                Section section = document.Sections[1];
                Borders borders = null;
                Border top = null;
                Border bottom = null;
                Border left = null;
                Border right = null;
                try
                {
                    borders = section.Borders;
                    int bordersShadow = -1;
                    try { dynamic db = borders; bordersShadow = Convert.ToInt32(db.Shadow); } catch { }
                    if (bordersShadow != 0 && bordersShadow != -1)
                        return Miss("上下のページ罫線が指定どおりではありません。");

                    top = borders[WdBorderType.wdBorderTop];
                    bottom = borders[WdBorderType.wdBorderBottom];
                    left = borders[WdBorderType.wdBorderLeft];
                    right = borders[WdBorderType.wdBorderRight];

                    if (!PageBorderSideOk(top) || !PageBorderSideOk(bottom))
                        return Miss("上下のページ罫線が指定どおりではありません。");

                    if (!PageBorderSideAbsent(left) || !PageBorderSideAbsent(right))
                        return Miss("上下のページ罫線が指定どおりではありません。");

                    if (!PageBorderTopBottomAccent5(document))
                        return Miss("上下のページ罫線が指定どおりではありません。");

                    return true;
                }
                finally
                {
                    if (right != null) Marshal.ReleaseComObject(right);
                    if (left != null) Marshal.ReleaseComObject(left);
                    if (bottom != null) Marshal.ReleaseComObject(bottom);
                    if (top != null) Marshal.ReleaseComObject(top);
                    if (borders != null) Marshal.ReleaseComObject(borders);
                    Marshal.ReleaseComObject(section);
                }
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool PageBorderSideOk(Border border)
        {
            if (border == null)
                return false;
            try
            {
                if ((WdLineStyle)border.LineStyle != WdLineStyle.wdLineStyleSingle)
                    return false;
                return (WdLineWidth)border.LineWidth == WdLineWidth.wdLineWidth150pt;
            }
            catch
            {
                return false;
            }
        }

        private static bool PageBorderSideAbsent(Border border)
        {
            if (border == null)
                return true;
            try
            {
                return !PageBorderSideOk(border);
            }
            catch
            {
                return true;
            }
        }

        private static bool PageBorderTopBottomAccent5(Document document)
        {
            string xml;
            try { xml = document.WordOpenXML ?? string.Empty; }
            catch { xml = string.Empty; }

            if (string.IsNullOrEmpty(xml))
                return false;

            if (!TryExtractPgBordersXml(xml, out string pgBorders))
                return false;

            if (!IsPgBorderSideAccent5WithWidth(pgBorders, "top"))
                return false;
            if (!IsPgBorderSideAccent5WithWidth(pgBorders, "bottom"))
                return false;

            if (PgBorderSidePresent(pgBorders, "left") && !PgBorderSideIsNone(pgBorders, "left"))
                return false;
            if (PgBorderSidePresent(pgBorders, "right") && !PgBorderSideIsNone(pgBorders, "right"))
                return false;

            return true;
        }

        private static bool TryExtractPgBordersXml(string xml, out string pgBordersXml)
        {
            pgBordersXml = null;
            int start = xml.IndexOf("<w:pgBorders", StringComparison.OrdinalIgnoreCase);
            if (start < 0)
                return false;
            int end = xml.IndexOf("</w:pgBorders>", start, StringComparison.OrdinalIgnoreCase);
            if (end < 0)
                return false;
            end += "</w:pgBorders>".Length;
            pgBordersXml = xml.Substring(start, end - start);
            return true;
        }

        private static bool IsPgBorderSideAccent5WithWidth(string pgBordersXml, string sideName)
        {
            var match = Regex.Match(pgBordersXml, $@"<w:{sideName}\s+([^/>]*)/>", RegexOptions.IgnoreCase);
            if (!match.Success)
                return false;
            string attrs = match.Groups[1].Value;
            if (!Regex.IsMatch(attrs, @"w:themeColor\s*=\s*""accent5""", RegexOptions.IgnoreCase))
                return false;
            if (!Regex.IsMatch(attrs, @"w:sz\s*=\s*""12""", RegexOptions.IgnoreCase))
            {
                if (!Regex.IsMatch(attrs, @"w:sz\s*=\s*""(1[01]|[8-9])""", RegexOptions.IgnoreCase))
                    return false;
            }
            return true;
        }

        private static bool PgBorderSidePresent(string pgBordersXml, string sideName)
        {
            return Regex.IsMatch(pgBordersXml, $@"<w:{sideName}\s", RegexOptions.IgnoreCase);
        }

        private static bool PgBorderSideIsNone(string pgBordersXml, string sideName)
        {
            var match = Regex.Match(pgBordersXml, $@"<w:{sideName}\s+([^/>]*)/>", RegexOptions.IgnoreCase);
            if (!match.Success)
                return true;
            string attrs = match.Groups[1].Value;
            return Regex.IsMatch(attrs, @"w:val\s*=\s*""none""", RegexOptions.IgnoreCase);
        }

        private bool CheckTask_1_4_02_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Paragraph heading = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                heading = FindParagraphContaining(document, "当日の流れ");
                if (heading == null)
                    return Miss("見出し「当日の流れ」がありません。");

                int afterStart = heading.Range.End;
                foreach (string prefix in FlowParagraphPrefixes)
                {
                    Paragraph target = FindParagraphStartingAfter(document, prefix, afterStart);
                    if (target == null)
                    {
                        Miss("「" + prefix + "…」の段落が見つかりません。");
                        return false;
                    }

                    try
                    {
                        if (!ParagraphHasLeftIndentChars(target, LeftIndentCharTarget))
                        {
                            Miss("「" + prefix + "…」などの段落の左インデントが3字になっていません。");
                            return false;
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(target);
                    }
                }

                return true;
            }
            finally
            {
                if (heading != null)
                    Marshal.ReleaseComObject(heading);
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool ParagraphHasLeftIndentChars(Paragraph paragraph, float targetChars)
        {
            try
            {
                dynamic fmt = paragraph.Format;
                try
                {
                    float units = (float)fmt.CharacterUnitLeftIndent;
                    return Math.Abs(units - targetChars) <= LeftIndentCharTolerance;
                }
                catch
                {
                    float leftPt = (float)paragraph.Format.LeftIndent;
                    float sizePt = 10.5f;
                    try { sizePt = (float)paragraph.Range.Font.Size; } catch { }
                    if (sizePt <= 0f)
                        sizePt = 10.5f;
                    float expectedPt = targetChars * sizePt;
                    return Math.Abs(leftPt - expectedPt) <= sizePt * 0.35f;
                }
            }
            catch
            {
                return false;
            }
        }

        private bool CheckTask_1_4_03_Impl(string filePath)
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

                const string titleText = "ジャズナイト開催計画";
                if (string.IsNullOrEmpty(xml) || xml.IndexOf(titleText, StringComparison.Ordinal) < 0)
                    return Miss("タイトル「ジャズナイト開催計画」がありません。");

                string scopeXml = ExtractTitleTextEffectScopeXml(xml, titleText);
                if (string.IsNullOrEmpty(scopeXml))
                    return Miss("「ジャズナイト開催計画」の文字の効果が指定どおりではありません。");

                if (!TitleTextEffectLooksCorrect(scopeXml))
                    return Miss("「ジャズナイト開催計画」の文字の効果が指定どおりではありません。");

                return true;
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        /// <summary>［塗りつぶし：青、アクセント1；影］相当（accent1 塗り + 影。輪郭付きプリセットは除外）。</summary>
        private static bool TitleTextEffectLooksCorrect(string scopeXml)
        {
            if (string.IsNullOrEmpty(scopeXml))
                return false;

            if (!XmlHasTextShadow(scopeXml))
                return false;

            if (!XmlHasAccent1TextFill(scopeXml))
                return false;

            if (XmlHasProminentTextOutline(scopeXml))
                return false;

            return true;
        }

        private static bool XmlHasTextShadow(string xml)
        {
            return xml.IndexOf("<w14:shadow", StringComparison.OrdinalIgnoreCase) >= 0
                || xml.IndexOf("<w15:shadow", StringComparison.OrdinalIgnoreCase) >= 0
                || xml.IndexOf("outerShdw", StringComparison.OrdinalIgnoreCase) >= 0
                || xml.IndexOf("innerShdw", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool XmlHasAccent1TextFill(string xml)
        {
            var fillMatch = Regex.Match(
                xml,
                @"<w14:textFill\b[^>]*>.*?</w14:textFill>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (fillMatch.Success)
            {
                string fillXml = fillMatch.Value;
                if (Regex.IsMatch(fillXml, @"accent1", RegexOptions.IgnoreCase))
                    return true;
            }

            if (Regex.IsMatch(xml, @"w14:schemeClr\s+w14:val=""accent1""", RegexOptions.IgnoreCase))
                return true;

            if (Regex.IsMatch(xml, @"<w:color\s+[^>]*w:themeColor=""accent1""", RegexOptions.IgnoreCase))
                return true;

            return Regex.IsMatch(
                xml,
                @"<w14:props\b[^>]*>.*?accent1.*?</w14:props>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
        }

        /// <summary>塗り+輪郭系プリセット（教材 P6 等）を除外。影のみプリセットは通す。</summary>
        private static bool XmlHasProminentTextOutline(string xml)
        {
            var outlineMatch = Regex.Match(
                xml,
                @"<w14:textOutline\b[^>]*>.*?</w14:textOutline>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (!outlineMatch.Success)
                return false;

            string outlineXml = outlineMatch.Value;
            if (!Regex.IsMatch(outlineXml, @"accent1", RegexOptions.IgnoreCase))
                return false;

            return Regex.IsMatch(outlineXml, @"w14:width\s+w14:val=""(?!0"")", RegexOptions.IgnoreCase);
        }

        private static string ExtractTitleTextEffectScopeXml(string xml, string titleText)
        {
            string paragraphXml = ExtractParagraphXmlContaining(xml, titleText);
            if (!string.IsNullOrEmpty(paragraphXml)
                && XmlHasTextShadow(paragraphXml)
                && XmlHasAccent1TextFill(paragraphXml))
            {
                return paragraphXml;
            }

            return ExtractXmlWindowContaining(xml, titleText, 14000, 14000);
        }

        private static string ExtractXmlWindowContaining(string xml, string text, int beforeChars, int afterChars)
        {
            int idx = xml.IndexOf(text, StringComparison.Ordinal);
            if (idx < 0)
                return null;

            int start = Math.Max(0, idx - beforeChars);
            int end = Math.Min(xml.Length, idx + text.Length + afterChars);
            return xml.Substring(start, end - start);
        }

        private static string ExtractParagraphXmlContaining(string xml, string text)
        {
            int idx = xml.IndexOf(text, StringComparison.Ordinal);
            if (idx < 0)
                return null;

            int pStart = xml.LastIndexOf("<w:p", idx, StringComparison.OrdinalIgnoreCase);
            if (pStart < 0)
                return null;

            int pEnd = xml.IndexOf("</w:p>", idx, StringComparison.OrdinalIgnoreCase);
            if (pEnd < 0)
                return null;
            pEnd += "</w:p>".Length;

            return xml.Substring(pStart, pEnd - pStart);
        }

        private bool CheckTask_1_4_04_Impl(string filePath)
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

                if (DocumentHasCommentOn19Opening(document))
                    return true;

                return Miss("「19時開場」に指定のコメントがありません。");
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool DocumentHasCommentOn19Opening(Document document)
        {
            const string anchorPhrase = "19時開場";
            const string bodyPhrase = "開場時刻を再確認";

            Comments comments = null;
            try
            {
                comments = document.Comments;
                int count = comments.Count;
                for (int i = 1; i <= count; i++)
                {
                    Comment comment = null;
                    Range scope = null;
                    try
                    {
                        comment = comments[i];
                        string body = NormalizeCommentBody(GetCommentBodyText(comment));
                        if (body.IndexOf(bodyPhrase, StringComparison.Ordinal) < 0)
                            continue;

                        scope = comment.Scope;
                        string scopeText = NormalizeCommentBody(scope?.Text ?? string.Empty);
                        if (scopeText.IndexOf(anchorPhrase, StringComparison.Ordinal) >= 0)
                            return true;

                        if (CommentAnchorOverlapsFind(document, comment, anchorPhrase))
                            return true;
                    }
                    catch
                    {
                        // モダンコメント等で COM が失敗する場合は次へ
                    }
                    finally
                    {
                        if (scope != null)
                            Marshal.ReleaseComObject(scope);
                        if (comment != null)
                            Marshal.ReleaseComObject(comment);
                    }
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                if (comments != null)
                    Marshal.ReleaseComObject(comments);
            }

            return false;
        }

        private static bool CommentAnchorOverlapsFind(Document document, Comment comment, string anchorPhrase)
        {
            Range findRange = null;
            Find find = null;
            Range commentRange = null;
            try
            {
                commentRange = comment.Range;
                int cStart = commentRange.Start;
                int cEnd = commentRange.End;

                findRange = document.Content;
                find = findRange.Find;
                find.ClearFormatting();
                find.Text = anchorPhrase;
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;

                while (find.Execute())
                {
                    int fStart = findRange.Start;
                    int fEnd = findRange.End;
                    if (RangesOverlap(cStart, cEnd, fStart, fEnd))
                        return true;
                    findRange.SetRange(fEnd, document.Content.End);
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                if (find != null)
                    Marshal.ReleaseComObject(find);
                if (findRange != null)
                    Marshal.ReleaseComObject(findRange);
                if (commentRange != null)
                    Marshal.ReleaseComObject(commentRange);
            }

            return false;
        }

        private static bool RangesOverlap(int aStart, int aEnd, int bStart, int bEnd)
        {
            return aStart < bEnd && bStart < aEnd;
        }

        private static string GetCommentBodyText(Comment comment)
        {
            try { return comment.Range?.Text ?? string.Empty; }
            catch { return string.Empty; }
        }

        private static string NormalizeCommentBody(string s)
        {
            if (string.IsNullOrEmpty(s))
                return string.Empty;
            return s.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\a", string.Empty).Trim();
        }

        private static Paragraph FindParagraphStartingAfter(Document document, string prefix, int minStart)
        {
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                try
                {
                    if (paragraph.Range.Start < minStart)
                        continue;

                    string text = ParagraphText(paragraph);
                    if (text.StartsWith(prefix, StringComparison.Ordinal))
                        return paragraph;
                }
                finally
                {
                    if (paragraph.Range.Start < minStart || !ParagraphText(paragraph).StartsWith(prefix, StringComparison.Ordinal))
                        Marshal.ReleaseComObject(paragraph);
                }
            }

            return null;
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

        private static string ParagraphText(Paragraph paragraph)
        {
            try
            {
                return (paragraph.Range.Text ?? string.Empty).Trim('\r', '\n', '\a', ' ', '　');
            }
            catch
            {
                return string.Empty;
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