using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Libraries;
using Microsoft.Office.Interop.Word;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題第1回 プロジェクト3。3-1〜3-4 は最終状態で判定する。
    /// 3-3 はドキュメント検査の操作自体は見ず、コメントが残っていないことのみ見る（教材4-7と同様）。
    /// </summary>
    public class WordChecker1_3_PV1
    {
        /// <summary>段組みダイアログで2段目16字にしたとき、1段目が約15.54字になる典型比。</summary>
        private const float SecondColumnCharCount = 16f;
        private const float FirstColumnReferenceCharCount = 15.54f;
        private const float ColumnWidthRatioTolerance = 0.05f;

        public bool CheckTask_1_3_01() => RunCheck(CheckTask_1_3_01_Impl);
        public bool CheckTask_1_3_02() => RunCheck(CheckTask_1_3_02_Impl);
        public bool CheckTask_1_3_03() => RunCheck(CheckTask_1_3_03_Impl);
        public bool CheckTask_1_3_04() => RunCheck(CheckTask_1_3_04_Impl);

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

        private bool CheckTask_1_3_01_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Paragraph startPara = null;
            Paragraph endPara = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                startPara = FindParagraphContaining(document, "見どころ");
                if (startPara == null)
                    return Miss("「見どころ」を含む段落がありません。");

                int rangeStart = startPara.Range.Start;
                endPara = FindParagraphContainingAfter(document, "水飲み場", rangeStart);
                if (endPara == null)
                    return Miss("「…水飲み場」までの範囲が確認できません。");

                Section startSection = startPara.Range.Sections[1];
                Section endSection = endPara.Range.Sections[1];
                try
                {
                    if (startSection.Index != endSection.Index)
                        return Miss("「見どころ」から「…水飲み場」までが同じ段組みになっていません。");

                    if (!ParagraphsThroughEndShareSection(document, rangeStart, endPara.Range.End, startSection.Index))
                        return Miss("「見どころ」から「…水飲み場」までが同じ段組みになっていません。");

                    TextColumns cols = startSection.PageSetup.TextColumns;
                    try
                    {
                        if (cols.Count != 2)
                            return Miss("選択範囲の段組みが2段になっていません。");

                        if (cols.EvenlySpaced != 0)
                            return Miss("「段の幅を同じにする」がオフになっていません。");

                        TextColumn c1 = cols[1];
                        TextColumn c2 = cols[2];
                        try
                        {
                            float w1 = (float)c1.Width;
                            float w2 = (float)c2.Width;
                            if (w1 <= 0f)
                                return Miss("2段目の幅が16字になっていません。");

                            float expectedRatio = SecondColumnCharCount / FirstColumnReferenceCharCount;
                            float actualRatio = w2 / w1;
                            if (Math.Abs(actualRatio - expectedRatio) > ColumnWidthRatioTolerance)
                                return Miss("2段目の幅が16字になっていません。");

                            return true;
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(c1);
                            Marshal.ReleaseComObject(c2);
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(cols);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(startSection);
                    Marshal.ReleaseComObject(endSection);
                }
            }
            finally
            {
                if (endPara != null)
                    Marshal.ReleaseComObject(endPara);
                if (startPara != null)
                    Marshal.ReleaseComObject(startPara);
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_3_02_Impl(string filePath)
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

                heading = FindParagraphByText(document, "3.夕暮れの道");
                if (heading == null)
                    return Miss("見出し「3.夕暮れの道」がありません。");

                if (!ParagraphStartsOnNewPage(document, heading))
                    return Miss("見出し「3.夕暮れの道」が次のページから始まっていません。");

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

        private bool CheckTask_1_3_03_Impl(string filePath)
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

                // 教材4-7と同様: 検査UIは見ず、削除結果（コメント無し）のみ。
                if (DocumentHasAnyComment(document))
                    return Miss("コメントが削除されていません。");

                string xml;
                try { xml = document.WordOpenXML ?? string.Empty; }
                catch { xml = string.Empty; }

                if (XmlHasAnyComment(xml))
                    return Miss("コメントが削除されていません。");

                return true;
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

        private static bool DocumentHasAnyComment(Document document)
        {
            Comments comments = null;
            try
            {
                comments = document.Comments;
                return comments != null && comments.Count > 0;
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
        }

        private static bool XmlHasAnyComment(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            foreach (Match part in Regex.Matches(
                xml,
                @"<pkg:part pkg:name=""/word/comments[^""]*""[^>]*>.*?</pkg:part>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                if (Regex.IsMatch(part.Value, @"<w:comment\b", RegexOptions.IgnoreCase))
                    return true;
            }

            return false;
        }

        private bool CheckTask_1_3_04_Impl(string filePath)
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

                if (document.ProtectionType != WdProtectionType.wdAllowOnlyRevisions)
                    return Miss("変更履歴の記録を解除できないロックになっていません。");

                bool tracking = false;
                try { tracking = document.TrackRevisions; } catch { }
                if (!tracking)
                    return Miss("変更履歴の記録がオンになっていません。");

                try
                {
                    document.Unprotect(string.Empty);
                    document.Protect(
                        WdProtectionType.wdAllowOnlyRevisions,
                        NoReset: true,
                        Password: string.Empty,
                        UseIRM: false,
                        EnforceStyleLock: false);
                    return true;
                }
                catch (COMException)
                {
                    return Miss("パスワードなしの変更履歴ロックになっていません。");
                }
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static Paragraph FindParagraphContainingAfter(Document document, string part, int minStart)
        {
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                try
                {
                    if (paragraph.Range.Start < minStart)
                        continue;

                    if (ParagraphContainsText(paragraph, part))
                        return paragraph;
                }
                finally
                {
                    if (!ParagraphContainsText(paragraph, part) || paragraph.Range.Start < minStart)
                        Marshal.ReleaseComObject(paragraph);
                }
            }

            return null;
        }

        private static bool ParagraphsThroughEndShareSection(Document document, int rangeStart, int rangeEnd, int sectionIndex)
        {
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                try
                {
                    int start = paragraph.Range.Start;
                    if (start < rangeStart)
                        continue;
                    if (start >= rangeEnd)
                        break;

                    Section section = paragraph.Range.Sections[1];
                    try
                    {
                        if (section.Index != sectionIndex)
                            return false;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(section);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(paragraph);
                }
            }

            return true;
        }

        private static bool ParagraphStartsOnNewPage(Document document, Paragraph heading)
        {
            try
            {
                int breakBefore = 0;
                try { breakBefore = heading.Format.PageBreakBefore; } catch { }
                if (breakBefore != 0)
                    return true;

                int start = heading.Range.Start;
                if (start > 0)
                {
                    Range before = null;
                    try
                    {
                        before = document.Range(start - 1, start);
                        string t = before.Text ?? string.Empty;
                        if (t.IndexOf('\f') >= 0)
                            return true;
                        if (t.Length > 0 && t[0] == (char)12)
                            return true;
                    }
                    finally
                    {
                        if (before != null)
                            Marshal.ReleaseComObject(before);
                    }
                }

                Paragraph prev = null;
                try
                {
                    prev = heading.Previous();
                    if (prev == null)
                    {
                        int page = (int)heading.Range.Information[WdInformation.wdActiveEndPageNumber];
                        return page >= 1;
                    }

                    int pageH = (int)heading.Range.Information[WdInformation.wdActiveEndPageNumber];
                    int pageP = (int)prev.Range.Information[WdInformation.wdActiveEndPageNumber];
                    return pageH > pageP;
                }
                finally
                {
                    if (prev != null)
                        Marshal.ReleaseComObject(prev);
                }
            }
            catch
            {
                return false;
            }
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

        private static Paragraph FindParagraphByText(Document document, string expected)
        {
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                if (ParagraphText(paragraph) == expected)
                    return paragraph;
                Marshal.ReleaseComObject(paragraph);
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

        private static bool Miss(string reason)
        {
            WordScoreExplanation.Note(reason);
            return false;
        }
    }
}