using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Libraries;
using Microsoft.Office.Interop.Word;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題第1回 プロジェクト5。5-1〜5-4 は最終状態のみ（ログ・破壊ゲートなし）。
    /// </summary>
    public class WordChecker1_5_PV1
    {
        private const char DiamondBullet = '\u25C6';
        private const int BodyPrefixScanLength = 16;
        private const float ColumnWidthTolerancePt = 1.5f;
        private const int GridColWidthToleranceTwips = 20;

        public bool CheckTask_1_5_01() => RunCheck(CheckTask_1_5_01_Impl);
        public bool CheckTask_1_5_02() => RunCheck(CheckTask_1_5_02_Impl);
        public bool CheckTask_1_5_03() => RunCheck(CheckTask_1_5_03_Impl);
        public bool CheckTask_1_5_04() => RunCheck(CheckTask_1_5_04_Impl);

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

        private bool CheckTask_1_5_01_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Paragraph heading = null;
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

                heading = FindParagraphContaining(document, "手入れの道具");
                if (heading == null)
                    return Miss("見出し「手入れの道具」がありません。");

                int afterHeading = heading.Range.End;
                startPara = FindParagraphBodyStartsWithAfter(document, "土", afterHeading);
                if (startPara == null)
                    return Miss("「土」から「支柱」までが◆の箇条書きになっていません。");

                endPara = FindParagraphBodyStartsWithAfter(document, "支柱", startPara.Range.Start);
                if (endPara == null)
                    endPara = FindParagraphContainingAfter(document, "支柱", startPara.Range.Start);
                if (endPara == null)
                    return Miss("「土」から「支柱」までが◆の箇条書きになっていません。");

                int rangeStart = startPara.Range.Start;
                int rangeEnd = endPara.Range.End;

                string xml;
                try { xml = document.WordOpenXML ?? string.Empty; }
                catch { xml = string.Empty; }

                if (!BulletedParagraphsInRangeHaveDiamondBullets(document, xml, rangeStart, rangeEnd))
                    return Miss("「土」から「支柱」までが◆の箇条書きになっていません。");

                return true;
            }
            finally
            {
                if (endPara != null)
                    Marshal.ReleaseComObject(endPara);
                if (startPara != null)
                    Marshal.ReleaseComObject(startPara);
                if (heading != null)
                    Marshal.ReleaseComObject(heading);
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool BulletedParagraphsInRangeHaveDiamondBullets(
            Document document,
            string documentXml,
            int rangeStart,
            int rangeEnd)
        {
            int bulletParagraphCount = 0;

            foreach (Paragraph paragraph in document.Paragraphs)
            {
                try
                {
                    int start = paragraph.Range.Start;
                    if (start < rangeStart)
                        continue;
                    if (start > rangeEnd)
                        break;

                    if (string.IsNullOrEmpty(ParagraphText(paragraph)))
                        continue;

                    if (!ParagraphIsBulletedList(paragraph))
                        continue;

                    bulletParagraphCount++;
                    if (!ParagraphHasDiamondBullet(paragraph, documentXml))
                        return false;
                }
                finally
                {
                    Marshal.ReleaseComObject(paragraph);
                }
            }

            return bulletParagraphCount > 0;
        }

        private static bool ParagraphIsBulletedList(Paragraph paragraph)
        {
            ListFormat lf = null;
            try
            {
                lf = paragraph.Range.ListFormat;
                WdListType listType = lf.ListType;
                return listType == WdListType.wdListBullet
                    || listType == WdListType.wdListPictureBullet;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (lf != null)
                    Marshal.ReleaseComObject(lf);
            }
        }

        private static bool ParagraphHasDiamondBullet(Paragraph paragraph, string documentXml)
        {
            ListFormat lf = null;
            try
            {
                lf = paragraph.Range.ListFormat;
                WdListType listType = lf.ListType;
                if (listType != WdListType.wdListBullet && listType != WdListType.wdListPictureBullet)
                    return false;

                string listString = string.Empty;
                try { listString = lf.ListString ?? string.Empty; } catch { }
                if (listString.IndexOf(DiamondBullet) >= 0 || listString.IndexOf('◆') >= 0)
                    return true;

                if (listType == WdListType.wdListPictureBullet)
                    return true;

                if (ListTemplateUsesDiamondSymbol(lf))
                    return true;

                if (ParagraphXmlIndicatesDiamondBullet(documentXml, paragraph))
                    return true;

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (lf != null)
                    Marshal.ReleaseComObject(lf);
            }
        }

        private static bool ParagraphXmlIndicatesDiamondBullet(string documentXml, Paragraph paragraph)
        {
            if (string.IsNullOrEmpty(documentXml))
                return false;

            string body = ParagraphBodyForMatch(paragraph);
            if (string.IsNullOrEmpty(body) || body.Length < 2)
                return false;

            string needle = body.Length > 12 ? body.Substring(0, 12) : body;
            int idx = documentXml.IndexOf(needle, StringComparison.Ordinal);
            if (idx < 0)
                return false;

            string paraXml = ExtractParagraphXmlContainingIndex(documentXml, idx);
            if (string.IsNullOrEmpty(paraXml))
                return false;

            if (paraXml.IndexOf("<w:numPr", StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            if (paraXml.IndexOf(DiamondBullet) >= 0 || paraXml.IndexOf('◆') >= 0)
                return true;

            if (Regex.IsMatch(paraXml, @"<w:sym\b", RegexOptions.IgnoreCase))
                return true;

            return true;
        }

        private static string ExtractParagraphXmlContainingIndex(string xml, int index)
        {
            int pStart = xml.LastIndexOf("<w:p", index, StringComparison.OrdinalIgnoreCase);
            if (pStart < 0)
                return null;

            int pEnd = xml.IndexOf("</w:p>", index, StringComparison.OrdinalIgnoreCase);
            if (pEnd < 0)
                return null;
            pEnd += "</w:p>".Length;

            return xml.Substring(pStart, pEnd - pStart);
        }

        private static string ParagraphBodyForMatch(Paragraph paragraph)
        {
            string text = ParagraphText(paragraph);
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            int i = 0;
            while (i < text.Length && i < BodyPrefixScanLength)
            {
                char c = text[i];
                if (c == '\t' || c == ' ' || c == '　' || char.IsControl(c))
                {
                    i++;
                    continue;
                }

                if (!char.IsLetterOrDigit(c) && c != '土' && c != '支')
                {
                    i++;
                    continue;
                }

                break;
            }

            return i > 0 && i < text.Length ? text.Substring(i) : text;
        }

        private static bool ParagraphBodyStartsWith(Paragraph paragraph, string prefix)
        {
            string body = ParagraphBodyForMatch(paragraph);
            if (string.IsNullOrEmpty(body))
                return false;

            if (body.StartsWith(prefix, StringComparison.Ordinal))
                return true;

            int idx = body.IndexOf(prefix, StringComparison.Ordinal);
            return idx >= 0 && idx <= 3;
        }

        private static bool ListTemplateUsesDiamondSymbol(ListFormat lf)
        {
            ListTemplate template = null;
            ListLevels levels = null;
            ListLevel level = null;
            try
            {
                template = lf.ListTemplate;
                if (template == null)
                    return false;

                levels = template.ListLevels;
                if (levels == null || levels.Count < 1)
                    return false;

                level = levels[1];
                string numberFormat = string.Empty;
                try { numberFormat = level.NumberFormat ?? string.Empty; } catch { }
                if (numberFormat.IndexOf(DiamondBullet) >= 0 || numberFormat.IndexOf('◆') >= 0)
                    return true;

                try
                {
                    WdListNumberStyle style = level.NumberStyle;
                    if (style == WdListNumberStyle.wdListNumberStyleBullet)
                    {
                        Font font = level.Font;
                        try
                        {
                            string name = font.Name ?? string.Empty;
                            if (name.IndexOf("Wingdings", StringComparison.OrdinalIgnoreCase) >= 0
                                || name.IndexOf("Symbol", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return true;
                            }
                        }
                        finally
                        {
                            if (font != null)
                                Marshal.ReleaseComObject(font);
                        }
                    }
                }
                catch { }

                return false;
            }
            finally
            {
                if (level != null)
                    Marshal.ReleaseComObject(level);
                if (levels != null)
                    Marshal.ReleaseComObject(levels);
                if (template != null)
                    Marshal.ReleaseComObject(template);
            }
        }

        private bool CheckTask_1_5_02_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Paragraph heading = null;
            Table table = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                heading = FindParagraphContaining(document, "今週の入荷");
                if (heading == null)
                    return Miss("見出し「今週の入荷」がありません。");

                table = FindFirstTableAfter(document, heading.Range.End);
                if (table == null)
                    return Miss("「今週の入荷」の表が「価格（円）」の昇順になっていません。");

                if (!TryFindHeaderColumn(table, "価格（円）", out int priceColumn))
                {
                    if (!TryFindHeaderColumn(table, "価格", out priceColumn))
                        return Miss("「今週の入荷」の表が「価格（円）」の昇順になっていません。");
                }

                var prices = new List<double>();
                for (int row = 2; row <= table.Rows.Count; row++)
                {
                    string text = CellText(table.Cell(row, priceColumn));
                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    if (!TryParseTableNumber(text, out double value))
                        return Miss("「今週の入荷」の表が「価格（円）」の昇順になっていません。");

                    prices.Add(value);
                }

                if (prices.Count < 2)
                    return Miss("「今週の入荷」の表が「価格（円）」の昇順になっていません。");

                for (int i = 1; i < prices.Count; i++)
                {
                    if (prices[i] + 0.0001 < prices[i - 1])
                        return Miss("「今週の入荷」の表が「価格（円）」の昇順になっていません。");
                }

                return true;
            }
            finally
            {
                if (table != null)
                    Marshal.ReleaseComObject(table);
                if (heading != null)
                    Marshal.ReleaseComObject(heading);
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_5_03_Impl(string filePath)
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

                if (LastDrawingDescriptionMatches(xml, "鉢植え"))
                    return true;

                if (LastInlinePictureDescriptionMatches(document, "鉢植え"))
                    return true;

                if (AnyPictureDescriptionMatches(document, xml, "鉢植え"))
                    return true;

                return Miss("「店頭の写真」の代替テキストが「鉢植え」になっていません。");
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool LastDrawingDescriptionMatches(string xml, string description)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            var matches = Regex.Matches(xml, @"<(?:wp|wp14):docPr\s+([^/>]*)/>", RegexOptions.IgnoreCase);
            if (matches.Count == 0)
                return false;

            string attrs = matches[matches.Count - 1].Groups[1].Value;
            return DescriptionAttributeMatches(attrs, description);
        }

        private static bool LastInlinePictureDescriptionMatches(Document document, string description)
        {
            int count = 0;
            try { count = document.InlineShapes.Count; } catch { }
            if (count < 1)
                return false;

            InlineShape shape = null;
            try
            {
                shape = document.InlineShapes[count];
                return AlternativeTextMatchesDescription(shape.AlternativeText, description);
            }
            catch
            {
                return false;
            }
            finally
            {
                if (shape != null)
                    Marshal.ReleaseComObject(shape);
            }
        }

        private static bool AnyPictureDescriptionMatches(Document document, string xml, string description)
        {
            int matchCount = 0;
            foreach (InlineShape shape in document.InlineShapes)
            {
                try
                {
                    if (AlternativeTextMatchesDescription(shape.AlternativeText, description))
                        matchCount++;
                }
                catch { }
                finally
                {
                    Marshal.ReleaseComObject(shape);
                }
            }

            foreach (Microsoft.Office.Interop.Word.Shape shape in document.Shapes)
            {
                try
                {
                    if (AlternativeTextMatchesDescription(shape.AlternativeText, description))
                        matchCount++;
                }
                catch { }
                finally
                {
                    Marshal.ReleaseComObject(shape);
                }
            }

            if (matchCount == 1)
                return true;

            if (matchCount > 1 && !string.IsNullOrEmpty(xml))
            {
                var docPrMatches = Regex.Matches(xml, @"<(?:wp|wp14):docPr\s+([^/>]*)/>", RegexOptions.IgnoreCase);
                int descrCount = 0;
                for (int i = 0; i < docPrMatches.Count; i++)
                {
                    if (DescriptionAttributeMatches(docPrMatches[i].Groups[1].Value, description))
                        descrCount++;
                }

                return descrCount == 1;
            }

            return false;
        }

        private static bool AlternativeTextMatchesDescription(string alternativeText, string description)
        {
            string alt = (alternativeText ?? string.Empty).Trim();
            return string.Equals(alt, description, StringComparison.Ordinal);
        }

        private static bool DescriptionAttributeMatches(string attrs, string description)
        {
            if (AttributeEquals(attrs, "descr", description))
                return true;
            return AttributeEquals(attrs, "description", description);
        }

        private static bool AttributeEquals(string attrs, string attrName, string expected)
        {
            var m = Regex.Match(
                attrs,
                attrName + @"\s*=\s*""([^""]*)""",
                RegexOptions.IgnoreCase);
            if (!m.Success)
                return false;

            return string.Equals(m.Groups[1].Value.Trim(), expected, StringComparison.Ordinal);
        }

        private bool CheckTask_1_5_04_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Paragraph heading = null;
            Table table = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                heading = FindParagraphContaining(document, "今週の入荷");
                if (heading == null)
                    return Miss("「入荷数」「単価」「金額」の列幅が揃っていません。");

                table = FindFirstTableAfter(document, heading.Range.End);
                if (table == null)
                    return Miss("「入荷数」「単価」「金額」の列幅が揃っていません。");

                if (!TryFindHeaderColumn(table, "入荷数", out int colA)
                    || !TryFindHeaderColumn(table, "単価", out int colB)
                    || !TryFindHeaderColumn(table, "金額", out int colC))
                {
                    return Miss("「入荷数」「単価」「金額」の列幅が揃っていません。");
                }

                if (ColumnsEqualWidthCom(table, colA, colB, colC))
                    return true;

                string xml;
                try { xml = document.WordOpenXML ?? string.Empty; }
                catch { xml = string.Empty; }

                if (!string.IsNullOrEmpty(xml)
                    && TableColumnsEqualWidthXml(xml, "今週の入荷", new[] { colA, colB, colC }))
                {
                    return true;
                }

                return Miss("「入荷数」「単価」「金額」の列幅が揃っていません。");
            }
            finally
            {
                if (table != null)
                    Marshal.ReleaseComObject(table);
                if (heading != null)
                    Marshal.ReleaseComObject(heading);
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool ColumnsEqualWidthCom(Table table, int colA, int colB, int colC)
        {
            float wA = GetColumnWidth(table, colA);
            float wB = GetColumnWidth(table, colB);
            float wC = GetColumnWidth(table, colC);
            if (wA <= 0f || wB <= 0f || wC <= 0f)
                return false;

            return Math.Abs(wA - wB) <= ColumnWidthTolerancePt
                && Math.Abs(wB - wC) <= ColumnWidthTolerancePt;
        }

        private static float GetColumnWidth(Table table, int columnIndex)
        {
            Column column = null;
            try
            {
                column = table.Columns[columnIndex];
                return (float)column.Width;
            }
            catch
            {
                return 0f;
            }
            finally
            {
                if (column != null)
                    Marshal.ReleaseComObject(column);
            }
        }

        private static bool TableColumnsEqualWidthXml(string xml, string headingText, int[] columnIndexes1Based)
        {
            string tableXml = ExtractTableXmlNearHeading(xml, headingText);
            if (string.IsNullOrEmpty(tableXml))
                return false;

            var gridCols = Regex.Matches(tableXml, @"<w:gridCol\s+w:w=""(\d+)""");
            if (gridCols.Count < columnIndexes1Based[columnIndexes1Based.Length - 1])
                return false;

            int firstWidth = ParseGridColWidth(gridCols, columnIndexes1Based[0]);
            if (firstWidth <= 0)
                return false;

            for (int i = 1; i < columnIndexes1Based.Length; i++)
            {
                int width = ParseGridColWidth(gridCols, columnIndexes1Based[i]);
                if (width <= 0 || Math.Abs(width - firstWidth) > GridColWidthToleranceTwips)
                    return false;
            }

            return true;
        }

        private static int ParseGridColWidth(MatchCollection gridCols, int columnIndex1Based)
        {
            int index = columnIndex1Based - 1;
            if (index < 0 || index >= gridCols.Count)
                return 0;
            return int.TryParse(gridCols[index].Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int w)
                ? w
                : 0;
        }

        private static string ExtractTableXmlNearHeading(string xml, string headingText)
        {
            int idx = xml.IndexOf(headingText, StringComparison.Ordinal);
            if (idx < 0)
                return null;

            int searchFrom = idx;
            int tblStart = xml.IndexOf("<w:tbl", searchFrom, StringComparison.OrdinalIgnoreCase);
            if (tblStart < 0)
                return null;

            int tblEnd = xml.IndexOf("</w:tbl>", tblStart, StringComparison.OrdinalIgnoreCase);
            if (tblEnd < 0)
                return null;
            tblEnd += "</w:tbl>".Length;

            return xml.Substring(tblStart, tblEnd - tblStart);
        }

        private static bool TryFindHeaderColumn(Table table, string headerPart, out int columnIndex)
        {
            columnIndex = 0;
            if (table.Rows.Count < 1)
                return false;

            int columnCount = table.Columns.Count;
            for (int c = 1; c <= columnCount; c++)
            {
                string text = CellText(table.Cell(1, c));
                if (text.IndexOf(headerPart, StringComparison.Ordinal) >= 0)
                {
                    columnIndex = c;
                    return true;
                }
            }

            return false;
        }

        private static bool TryParseTableNumber(string text, out double value)
        {
            value = 0d;
            text = (text ?? string.Empty).Trim();
            text = text.Replace(",", string.Empty).Replace("，", string.Empty);
            text = text.Replace("円", string.Empty).Trim();
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }

        private static Paragraph FindParagraphBodyStartsWithAfter(Document document, string prefix, int minStart)
        {
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                bool match = false;
                try
                {
                    if (paragraph.Range.Start < minStart)
                        continue;

                    match = ParagraphBodyStartsWith(paragraph, prefix);
                    if (match)
                        return paragraph;
                }
                finally
                {
                    if (!match)
                        Marshal.ReleaseComObject(paragraph);
                }
            }

            return null;
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
                    if (paragraph.Range.Start < minStart || !ParagraphContainsText(paragraph, part))
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