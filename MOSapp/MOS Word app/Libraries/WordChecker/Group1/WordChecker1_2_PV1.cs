using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Libraries;
using Microsoft.Office.Interop.Word;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題第1回 プロジェクト2。2-1〜2-5 は文書の最終状態で判定する。
    /// </summary>
    public class WordChecker1_2_PV1
    {
        private const float TargetImageHeightMm = 40f;
        private const float HeightTolerancePt = 2f;
        private const string DocumentTitle = "秋のパン教室";

        private static readonly string[] FormatTargetHeadings =
        {
            "パンの種類",
            "講師紹介",
            "申込方法"
        };

        private static readonly string[][] ScheduleRows =
        {
            new[] { "回", "日付" },
            new[] { "第1回", "10月4日" },
            new[] { "第2回", "10月18日" }
        };

        /// <summary>
        /// 目次をすべて更新したあとの期待エントリ（タイトル, レベル1or2, ページ）。
        /// 初期状態は「講座の趣旨」「一日の流れ」のみでページも古い。
        /// </summary>
        private static readonly TocExpectation[] ExpectedTocEntries =
        {
            new TocExpectation("講座の趣旨", 1, 1),
            new TocExpectation("店の経緯", 2, 1),
            new TocExpectation("窯", 2, 1),
            new TocExpectation("一日の流れ", 1, 2),
            new TocExpectation("持ち物", 1, 3),
            new TocExpectation("開講日", 2, 4),
            new TocExpectation("持参する道具", 2, 4)
        };

        private struct TocExpectation
        {
            public readonly string Title;
            public readonly int Level;
            public readonly int Page;

            public TocExpectation(string title, int level, int page)
            {
                Title = title;
                Level = level;
                Page = page;
            }
        }

        public bool CheckTask_1_2_01() => RunCheck(CheckTask_1_2_01_Impl);
        public bool CheckTask_1_2_02() => RunCheck(CheckTask_1_2_02_Impl);
        public bool CheckTask_1_2_03() => RunCheck(CheckTask_1_2_03_Impl);
        public bool CheckTask_1_2_04() => RunCheck(CheckTask_1_2_04_Impl);
        public bool CheckTask_1_2_05() => RunCheck(CheckTask_1_2_05_Impl);

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

        private bool CheckTask_1_2_01_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Paragraph source = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                source = FindParagraphByText(document, "教室の概要");
                if (source == null)
                    return Miss("見出し「教室の概要」がありません。");

                string sourceFp = GetParagraphFormatFingerprint(source);
                foreach (string heading in FormatTargetHeadings)
                {
                    Paragraph target = FindParagraphByText(document, heading);
                    if (target == null)
                    {
                        Miss("「" + heading + "」がありません。");
                        return false;
                    }

                    try
                    {
                        if (!string.Equals(sourceFp, GetParagraphFormatFingerprint(target), StringComparison.Ordinal))
                        {
                            Miss("「" + heading + "」の書式が「教室の概要」と一致していません。");
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
                if (source != null)
                    Marshal.ReleaseComObject(source);
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_2_02_Impl(string filePath)
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

                Range email = FindFirstTextRange(document, "bread@example.jp");
                if (email == null)
                    return Miss("「bread@example.jp」が見つかりません。");

                try
                {
                    if (email.Start <= 0)
                        return Miss("メールアドレスの前に記号がありません。");

                    if (EmailHasWingdingsCode46Before(document, email.Start))
                        return true;

                    return Miss("メールアドレスの前に Wingdings の文字コード46の記号がありません。");
                }
                finally
                {
                    Marshal.ReleaseComObject(email);
                }
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_2_03_Impl(string filePath)
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

                heading = FindParagraphByText(document, "開講日");
                if (heading == null)
                    return Miss("「開講日」がありません。");

                int headingEnd = heading.Range.End;
                Table table = FindFirstTableAfter(document, headingEnd);
                if (table == null)
                    return Miss("「開講日」の下に表がありません。");

                try
                {
                    if (table.Rows.Count != 3 || table.Columns.Count != 2)
                        return Miss("「開講日」の下の表が3行2列になっていません。");

                    for (int r = 1; r <= ScheduleRows.Length; r++)
                    {
                        for (int c = 1; c <= 2; c++)
                        {
                            string actual = CellText(table.Cell(r, c));
                            if (!string.Equals(actual, ScheduleRows[r - 1][c - 1], StringComparison.Ordinal))
                                return Miss("「開講日」の下の表の内容が合っていません。");
                        }
                    }

                    return true;
                }
                finally
                {
                    Marshal.ReleaseComObject(table);
                }
            }
            finally
            {
                if (heading != null)
                    Marshal.ReleaseComObject(heading);
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_2_04_Impl(string filePath)
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

                heading = FindParagraphByText(document, "持参する道具");
                if (heading == null)
                    return Miss("「持参する道具」がありません。");

                int after = heading.Range.End;
                Marshal.ReleaseComObject(heading);
                heading = null;

                Paragraph imagePara = FindNextParagraphWithInlinePicture(document, after, out InlineShape shape);
                if (imagePara == null || shape == null)
                    return Miss("「持参する道具」の下に画像がありません。");

                try
                {
                    float targetPt = TargetImageHeightMm * 72f / 25.4f;
                    float heightPt = 0f;
                    try { heightPt = shape.Height; } catch { }

                    if (Math.Abs(heightPt - targetPt) > HeightTolerancePt)
                        return Miss("画像の高さが40mmになっていません。");

                    WdParagraphAlignment align = WdParagraphAlignment.wdAlignParagraphLeft;
                    try { align = imagePara.Alignment; } catch { }
                    if (align != WdParagraphAlignment.wdAlignParagraphCenter)
                        return Miss("画像の段落が中央揃えになっていません。");

                    if (imagePara.Range.InlineShapes.Count < 1)
                        return Miss("画像が行内で挿入されていません。");

                    return true;
                }
                finally
                {
                    if (shape != null)
                        Marshal.ReleaseComObject(shape);
                    if (imagePara != null)
                        Marshal.ReleaseComObject(imagePara);
                }
            }
            finally
            {
                if (heading != null)
                    Marshal.ReleaseComObject(heading);
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_2_05_Impl(string filePath)
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

                Paragraph title = FindParagraphByText(document, DocumentTitle);
                if (title == null)
                    return Miss("タイトル「秋のパン教室」がありません。");

                int titleEnd = title.Range.End;
                Marshal.ReleaseComObject(title);

                Range tocRange = FindTocRangeAfter(document, titleEnd);
                if (tocRange == null)
                    return Miss("「秋のパン教室」の下に目次がありません。");

                try
                {
                    List<ParsedTocEntry> entries = ParseTocEntries(tocRange);
                    if (entries.Count == 0)
                        return Miss("目次の内容を読み取れません。");

                    if (!TocMatchesExpectation(entries, out string failReason))
                        return Miss(failReason);

                    return true;
                }
                finally
                {
                    Marshal.ReleaseComObject(tocRange);
                }
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

        private struct ParsedTocEntry
        {
            public string Title;
            public int Page;
            public int Level;
        }

        private static Range FindTocRangeAfter(Document document, int afterPosition)
        {
            try
            {
                TablesOfContents tocs = document.TablesOfContents;
                if (tocs != null && tocs.Count >= 1)
                {
                    TableOfContents best = null;
                    int bestStart = int.MaxValue;
                    for (int i = 1; i <= tocs.Count; i++)
                    {
                        TableOfContents toc = tocs[i];
                        try
                        {
                            int start = toc.Range.Start;
                            if (start >= afterPosition && start < bestStart)
                            {
                                if (best != null)
                                    Marshal.ReleaseComObject(best);
                                best = toc;
                                bestStart = start;
                            }
                            else
                            {
                                Marshal.ReleaseComObject(toc);
                            }
                        }
                        catch
                        {
                            Marshal.ReleaseComObject(toc);
                        }
                    }

                    if (best != null)
                    {
                        Range range = best.Range.Duplicate;
                        Marshal.ReleaseComObject(best);
                        return range;
                    }
                }
            }
            catch { }

            try
            {
                foreach (Field field in document.Fields)
                {
                    try
                    {
                        if (field.Type != WdFieldType.wdFieldTOC)
                            continue;
                        int start = field.Result.Start;
                        if (start < afterPosition)
                            continue;
                        return field.Result.Duplicate;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(field);
                    }
                }
            }
            catch { }

            return null;
        }

        private static List<ParsedTocEntry> ParseTocEntries(Range tocRange)
        {
            var entries = new List<ParsedTocEntry>();
            foreach (Paragraph paragraph in tocRange.Paragraphs)
            {
                try
                {
                    string text = ParagraphText(paragraph);
                    if (string.IsNullOrEmpty(text))
                        continue;
                    if (!TryParseTocLine(text, out string title, out int page))
                        continue;

                    int level = InferTocLevel(paragraph);
                    entries.Add(new ParsedTocEntry
                    {
                        Title = title,
                        Page = page,
                        Level = level
                    });
                }
                finally
                {
                    Marshal.ReleaseComObject(paragraph);
                }
            }

            return entries;
        }

        private static bool TryParseTocLine(string text, out string title, out int page)
        {
            title = null;
            page = 0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            string normalized = text
                .Replace('\t', ' ')
                .Replace('\u00A0', ' ')
                .Trim();
            Match match = Regex.Match(
                normalized,
                @"^(?<title>.+?)(?:[\.．…・\s]+)(?<page>\d+)\s*$");
            if (!match.Success)
                return false;

            title = match.Groups["title"].Value.Trim(' ', '　', '.', '．', '…', '・');
            if (string.IsNullOrEmpty(title))
                return false;
            if (!int.TryParse(match.Groups["page"].Value, out page) || page <= 0)
                return false;
            return true;
        }

        private static int InferTocLevel(Paragraph paragraph)
        {
            try
            {
                Style style = paragraph.get_Style() as Style;
                if (style != null)
                {
                    string name = "";
                    try { name = style.NameLocal ?? ""; } catch { }
                    Marshal.ReleaseComObject(style);
                    if (!string.IsNullOrEmpty(name))
                    {
                        if (Regex.IsMatch(name, @"(目次|TOC)\s*2", RegexOptions.IgnoreCase))
                            return 2;
                        if (Regex.IsMatch(name, @"(目次|TOC)\s*1", RegexOptions.IgnoreCase))
                            return 1;
                    }
                }
            }
            catch { }

            try
            {
                float left = (float)paragraph.Format.LeftIndent;
                if (left >= 10f)
                    return 2;
            }
            catch { }

            return 1;
        }

        private static bool TocMatchesExpectation(List<ParsedTocEntry> entries, out string failReason)
        {
            failReason = null;
            if (entries.Count < ExpectedTocEntries.Length)
            {
                failReason = "目次が更新されていません（項目が不足しています）。";
                return false;
            }

            for (int i = 0; i < ExpectedTocEntries.Length; i++)
            {
                TocExpectation expected = ExpectedTocEntries[i];
                ParsedTocEntry actual = entries[i];

                if (!string.Equals(actual.Title, expected.Title, StringComparison.Ordinal))
                {
                    failReason = "目次の項目順または内容が正しくありません。";
                    return false;
                }

                if (actual.Page != expected.Page)
                {
                    failReason = "目次「" + expected.Title + "」のページ番号が正しくありません。";
                    return false;
                }

                if (actual.Level != expected.Level)
                {
                    failReason = "目次「" + expected.Title + "」のレベルが正しくありません。";
                    return false;
                }
            }

            return true;
        }

        private static string GetParagraphFormatFingerprint(Paragraph paragraph)
        {
            string style = "";
            try
            {
                Style styleObj = paragraph.get_Style() as Style;
                if (styleObj != null)
                {
                    try { style = styleObj.NameLocal ?? ""; } catch { }
                    Marshal.ReleaseComObject(styleObj);
                }
            }
            catch { }

            float size = 0f;
            int bold = 0;
            int italic = 0;
            string fontName = "";
            Range r = null;
            try
            {
                r = paragraph.Range.Duplicate;
                if (r.Characters.Count >= 1)
                {
                    Range ch = r.Characters[1];
                    try
                    {
                        Font f = ch.Font;
                        try { size = f.Size; } catch { }
                        try { bold = f.Bold; } catch { }
                        try { italic = f.Italic; } catch { }
                        try { fontName = f.Name ?? ""; } catch { }
                        Marshal.ReleaseComObject(f);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(ch);
                    }
                }
            }
            catch { }
            finally
            {
                if (r != null)
                    Marshal.ReleaseComObject(r);
            }

            return style + "|" + size.ToString("0.##") + "|" + bold + "|" + italic + "|" + fontName;
        }

        private static bool EmailHasWingdingsCode46Before(Document document, int emailStart)
        {
            Paragraph para = null;
            try
            {
                para = document.Range(emailStart, emailStart).Paragraphs[1];
                int paraStart = para.Range.Start;
                int scanEnd = emailStart;
                if (scanEnd <= paraStart)
                    return false;

                if (RangeOpenXmlHasWingdingsCode46(document, paraStart, scanEnd))
                    return true;

                int pos = emailStart;
                int checkedCount = 0;
                const int maxLookback = 8;
                while (pos > paraStart && checkedCount < maxLookback)
                {
                    pos--;
                    Range one = null;
                    try
                    {
                        one = document.Range(pos, pos + 1);
                        string t = (one.Text ?? string.Empty).Replace("\r", "").Replace("\a", "");
                        if (t.Length == 0 || t == " " || t == "　")
                            continue;

                        string fontName = "";
                        try { fontName = one.Font?.Name ?? ""; } catch { }
                        if (fontName.IndexOf("Wingdings", StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;

                        checkedCount++;
                    }
                    finally
                    {
                        if (one != null)
                            Marshal.ReleaseComObject(one);
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
                if (para != null)
                    Marshal.ReleaseComObject(para);
            }
        }

        private static bool RangeOpenXmlHasWingdingsCode46(Document document, int start, int end)
        {
            try
            {
                if (end <= start)
                    return false;

                string snippet = document.Range(start, end).WordOpenXML;
                if (string.IsNullOrEmpty(snippet))
                    return false;

                if (Regex.IsMatch(
                        snippet,
                        @"<w:sym\b[^>]*w:font=""Wingdings""[^>]*w:char=""(?:002[Ee]|2[Ee]|0046|46|F02[Ee]|F046)""",
                        RegexOptions.IgnoreCase))
                {
                    return true;
                }

                return Regex.IsMatch(
                    snippet,
                    @"<w:rFonts\b[^>]*w:ascii=""Wingdings""",
                    RegexOptions.IgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static Paragraph FindNextParagraphWithInlinePicture(
            Document document,
            int afterPosition,
            out InlineShape pictureShape)
        {
            pictureShape = null;
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                try
                {
                    if (paragraph.Range.End <= afterPosition)
                        continue;

                    if (paragraph.Range.InlineShapes.Count < 1)
                        continue;

                    for (int i = 1; i <= paragraph.Range.InlineShapes.Count; i++)
                    {
                        InlineShape sh = paragraph.Range.InlineShapes[i];
                        try
                        {
                            if (IsInlinePictureShape(sh))
                            {
                                pictureShape = sh;
                                return paragraph;
                            }
                        }
                        finally
                        {
                            if (!ReferenceEquals(sh, pictureShape))
                                Marshal.ReleaseComObject(sh);
                        }
                    }
                }
                finally
                {
                    if (pictureShape == null)
                        Marshal.ReleaseComObject(paragraph);
                }
            }

            return null;
        }

        private static bool IsInlinePictureShape(InlineShape shape)
        {
            try
            {
                WdInlineShapeType t = shape.Type;
                if (t == WdInlineShapeType.wdInlineShapePicture ||
                    t == WdInlineShapeType.wdInlineShapeLinkedPicture)
                    return true;
            }
            catch { }

            try
            {
                string link = "";
                try { link = shape.LinkFormat?.SourceFullName ?? ""; } catch { }
                if (link.IndexOf("oven.png", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            catch { }

            try
            {
                string alt = shape.AlternativeText ?? "";
                if (alt.IndexOf("oven", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            catch { }

            return false;
        }

        private static Range FindFirstTextRange(Document document, string text)
        {
            Range search = null;
            try
            {
                search = document.Content.Duplicate;
                Find find = search.Find;
                find.ClearFormatting();
                find.Text = text;
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                if (!find.Execute())
                    return null;
                return search;
            }
            catch
            {
                if (search != null)
                    Marshal.ReleaseComObject(search);
                return null;
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

        private static Table FindFirstTableAfter(Document document, int headingEnd)
        {
            Table best = null;
            int bestStart = int.MaxValue;
            foreach (Table table in document.Tables)
            {
                int start = table.Range.Start;
                if (start >= headingEnd && start < bestStart)
                {
                    if (best != null)
                        Marshal.ReleaseComObject(best);
                    best = table;
                    bestStart = start;
                }
                else
                {
                    Marshal.ReleaseComObject(table);
                }
            }

            return best;
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

        private static string CellText(Cell cell)
        {
            try
            {
                return (cell.Range.Text ?? string.Empty).Trim('\r', '\n', '\a', ' ', '　');
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