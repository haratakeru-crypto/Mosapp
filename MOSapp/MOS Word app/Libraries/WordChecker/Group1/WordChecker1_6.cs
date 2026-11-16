using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_6
    {
        public bool CheckTask_1_6_01() => RunCheck(CheckTask_1_6_01_Impl);
        public bool CheckTask_1_6_02() => RunCheck(CheckTask_1_6_02_Impl);
        public bool CheckTask_1_6_03() => RunCheck(CheckTask_1_6_03_Impl);
        public bool CheckTask_1_6_04() => RunCheck(CheckTask_1_6_04_Impl);
        public bool CheckTask_1_6_05() => RunCheck(CheckTask_1_6_05_Impl);
        public bool CheckTask_1_6_06() => RunCheck(CheckTask_1_6_06_Impl);
        public bool CheckTask_1_6_07() => RunCheck(CheckTask_1_6_07_Impl);

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
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                // 1. ターゲット特定: 「月別催事内容」の直後にある表を抽出
                string tableXml = GetTableXmlNearText(xml, "月別催事内容");
                if (string.IsNullOrEmpty(tableXml))
                    return Miss("見出し「（2）月別催事内容」の下の表が見つかりません。");

                // 2. 構造チェック
                int rowCount = System.Text.RegularExpressions.Regex.Matches(tableXml, "<w:tr[ >]").Count;
                var firstRowMatch = System.Text.RegularExpressions.Regex.Match(tableXml, @"<w:tr[ >].*?</w:tr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                int colCount = firstRowMatch.Success ? System.Text.RegularExpressions.Regex.Matches(firstRowMatch.Value, "<w:tc[ >]").Count : 0;

                // 「文字列の幅に合わせる」設定の確認
                var tblPrMatch = System.Text.RegularExpressions.Regex.Match(tableXml, @"<w:tblPr\b[^>]*>.*?</w:tblPr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                string tblPrXml = tblPrMatch.Success ? tblPrMatch.Value : "";
                var tcPrMatch = System.Text.RegularExpressions.Regex.Match(tableXml, @"<w:tcPr\b[^>]*>.*?</w:tcPr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                string tcPrXml = tcPrMatch.Success ? tcPrMatch.Value : "";

                bool isAutoFit = tblPrXml.Contains(@"w:type=""auto""") && tcPrXml.Contains(@"w:type=""auto""");

                // 3. 分割状態の確認（6-2の影響を考慮）
                int table1EndIndex = xml.IndexOf(tableXml) + tableXml.Length;
                string followingXml = xml.Substring(table1EndIndex);
                var nextTableMatch = System.Text.RegularExpressions.Regex.Match(followingXml, @"<w:tbl\b[^>]*>.*?</w:tbl>", System.Text.RegularExpressions.RegexOptions.Singleline);
                bool hasSplitTable = nextTableMatch.Success && nextTableMatch.Index < 1000;

                // 行数判定: 「そのまま7行」または「分割されて1つ目が4行（かつ直後に表あり）」
                bool rowCountOk = (rowCount == 7) || (rowCount == 4 && hasSplitTable);

                // 4. 内容チェック: 分割されている場合は2つ目の表も対象に含める
                string combinedContentXml = tableXml + (hasSplitTable ? nextTableMatch.Value : "");
                bool contentOk = combinedContentXml.Contains("実施月") && combinedContentXml.Contains("富士の水だより");

                // 5. ログチェック
                bool logOk = LogReader.HasTaskEvidence(6, 1, "TableConvertTextToTable");

                if (logOk && rowCountOk && colCount == 2 && contentOk && isAutoFit)
                    return true;

                bool visualOk = rowCountOk && colCount == 2 && contentOk && isAutoFit;
                bool noted = false;
                if (!contentOk)
                {
                    WordScoreExplanation.Note("表に「実施月」から「富士の水だより」までの内容がありません。");
                    noted = true;
                }
                if (!rowCountOk)
                {
                    WordScoreExplanation.Note("表の行数が合っていません。");
                    noted = true;
                }
                if (colCount != 2)
                {
                    WordScoreExplanation.Note("表が2列になっていません。");
                    noted = true;
                }
                if (!isAutoFit)
                {
                    WordScoreExplanation.Note("表が文字列の幅に合わせた設定になっていません。");
                    noted = true;
                }
                // 操作記録は見た目がすべて合っているときだけ出す（4-6型）
                if (visualOk && !logOk)
                {
                    WordScoreExplanation.Note("文字列を表に変換した記録がありません。");
                    noted = true;
                }
                if (noted)
                    return false;
                return Miss("表への変換が指定どおりになっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_02_Impl(string filePath)
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

                string table1Xml = GetTableXmlNearText(xml, "月別催事内容");
                if (string.IsNullOrEmpty(table1Xml))
                    return Miss("見出し「（2）月別催事内容」の下の表が見つかりません。");

                int rowCount1 = System.Text.RegularExpressions.Regex.Matches(table1Xml, "<w:tr[ >]").Count;

                int table1EndIndex = xml.IndexOf(table1Xml) + table1Xml.Length;
                string followingXml = xml.Substring(table1EndIndex);
                var nextTableMatch = System.Text.RegularExpressions.Regex.Match(followingXml, @"<w:tbl\b[^>]*>.*?</w:tbl>", System.Text.RegularExpressions.RegexOptions.Singleline);
                bool hasSplitTable = nextTableMatch.Success && nextTableMatch.Index < 1000;

                if (rowCount1 == 4 && hasSplitTable)
                    return true;

                if (!hasSplitTable)
                    return Miss("表が分割されていません。");
                // 1つ目が N 行 → (N+1) 行目からの分割
                int splitFromRow = rowCount1 + 1;
                return Miss($"表が{splitFromRow}行目から分割されています。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_03_Impl(string filePath)
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

                var firstTableMatch = System.Text.RegularExpressions.Regex.Match(
                    xml, @"<w:tbl\b[^>]*>.*?</w:tbl>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!firstTableMatch.Success)
                    return Miss("1つ目の表が見つかりません。");
                string tableXml = firstTableMatch.Value;

                var allRows = System.Text.RegularExpressions.Regex.Matches(
                    tableXml, @"<w:tr[ >].*?</w:tr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (allRows.Count < 2)
                    return Miss("1つ目の表の構造が正しくありません。");

                int firstRowCellCount = System.Text.RegularExpressions.Regex.Matches(allRows[0].Value, "<w:tc[ >]").Count;
                bool firstRowOk = firstRowCellCount == 4;

                int wrongOtherRowIndex = -1; // 1-based 行番号
                int wrongOtherRowCells = 0;
                for (int i = 1; i < allRows.Count; i++)
                {
                    int cells = System.Text.RegularExpressions.Regex.Matches(allRows[i].Value, "<w:tc[ >]").Count;
                    if (cells != 2)
                    {
                        wrongOtherRowIndex = i + 1;
                        wrongOtherRowCells = cells;
                        break;
                    }
                }
                bool otherRowsOk = wrongOtherRowIndex < 0;

                if (firstRowOk && otherRowsOk)
                    return true;

                bool noted = false;
                if (!firstRowOk)
                {
                    // 元が2列で2列目だけ分割した前提: 全体セル数−1 が2列目の分割数
                    if (firstRowCellCount <= 2)
                        WordScoreExplanation.Note("1行目2列目が分割できていません。");
                    else
                        WordScoreExplanation.Note($"1行目2列目が{firstRowCellCount - 1}列に分割されています。");
                    noted = true;
                }
                if (!otherRowsOk)
                {
                    // 他行も同様に「○行目2列目を何列に分けたか」で伝える（まとめて分割した誤操作向け）
                    if (wrongOtherRowCells <= 2)
                        WordScoreExplanation.Note($"{wrongOtherRowIndex}行目のセル構成が変わっています。");
                    else
                        WordScoreExplanation.Note($"{wrongOtherRowIndex}行目2列目が{wrongOtherRowCells - 1}列に分割されています。");
                    noted = true;
                }
                if (noted)
                    return false;
                return Miss("1つ目の表のセル分割が指定どおりになっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_04_Impl(string filePath)
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

                string tableXml = GetTableXmlNearText(xml, "北海道のうまいもの市場売上");
                if (string.IsNullOrEmpty(tableXml))
                    return Miss("見出し「2.北海道のうまいもの市場売上」の下の表が見つかりません。");

                var firstRowMatch = System.Text.RegularExpressions.Regex.Match(
                    tableXml, @"<w:tr[ >].*?</w:tr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!firstRowMatch.Success)
                    return Miss("表の1行目が見つかりません。");
                string rowXml = firstRowMatch.Value;

                var outlineMatch = System.Text.RegularExpressions.Regex.Match(
                    rowXml, @"<w14:textOutline\b[^>]*>.*?</w14:textOutline>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                bool hasOutline = outlineMatch.Success;
                string outlineXml = hasOutline ? outlineMatch.Value : "";

                bool outlineIsOrange = hasOutline && (
                    System.Text.RegularExpressions.Regex.IsMatch(outlineXml, @"w14:schemeClr\s+w14:val=""accent2""")
                    || System.Text.RegularExpressions.Regex.IsMatch(
                        outlineXml, @"w14:srgbClr\s+w14:val=""ED7D31""",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase));

                bool fillIsOrange;
                var fillMatch = System.Text.RegularExpressions.Regex.Match(
                    rowXml, @"<w14:textFill\b[^>]*>.*?</w14:textFill>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (fillMatch.Success)
                {
                    string fillXml = fillMatch.Value;
                    fillIsOrange =
                        System.Text.RegularExpressions.Regex.IsMatch(fillXml, @"w14:schemeClr\s+w14:val=""accent2""")
                        || System.Text.RegularExpressions.Regex.IsMatch(
                            fillXml, @"w14:srgbClr\s+w14:val=""ED7D31""",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                }
                else
                {
                    fillIsOrange =
                        System.Text.RegularExpressions.Regex.IsMatch(rowXml, @"<w:color\s+[^>]*w:themeColor=""accent2""")
                        || System.Text.RegularExpressions.Regex.IsMatch(
                            rowXml, @"<w:color\s+[^>]*w:val=""ED7D31""",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                }

                bool hasAdditionalEffects =
                    rowXml.Contains("<w14:shadow")
                    || rowXml.Contains("<w14:reflection")
                    || rowXml.Contains("<w14:glow");

                if (outlineIsOrange && fillIsOrange && !hasAdditionalEffects)
                    return true;

                return Miss("表の1行目に「塗りつぶし：オレンジ、アクセント カラー2；輪郭：オレンジ アクセント カラー2」の文字の効果が設定されていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_05_Impl(string filePath)
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

                string tableXml = GetTableXmlNearText(xml, "北海道のうまいもの市場売上");
                if (string.IsNullOrEmpty(tableXml))
                    return Miss("見出し「2.北海道のうまいもの市場売上」の下の表が見つかりません。");

                var tblGridMatch = System.Text.RegularExpressions.Regex.Match(
                    tableXml, @"<w:tblGrid\b[^>]*>.*?</w:tblGrid>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!tblGridMatch.Success)
                    return Miss("表の列幅を確認できません。");
                string tblGridXml = tblGridMatch.Value;

                var gridCols = System.Text.RegularExpressions.Regex.Matches(
                    tblGridXml, @"<w:gridCol\s+w:w=""(\d+)""");
                if (gridCols.Count < 2)
                    return Miss("表の列幅を確認できません。");

                int firstWidth;
                if (!int.TryParse(gridCols[0].Groups[1].Value, out firstWidth))
                    return Miss("表の列幅を確認できません。");
                const int tolerance = 20;
                bool widthsEqual = true;
                for (int i = 1; i < gridCols.Count; i++)
                {
                    int width;
                    if (!int.TryParse(gridCols[i].Groups[1].Value, out width))
                        return Miss("表の列幅を確認できません。");
                    if (Math.Abs(width - firstWidth) > tolerance)
                    {
                        widthsEqual = false;
                        break;
                    }
                }

                // 既存どおり証跡キーは (6,4)。アドイン側の記録タスク番号に合わせる。
                bool logOk = LogReader.HasTaskEvidence(6, 4, "TableColumnsDistribute");

                if (widthsEqual && logOk)
                    return true;

                bool noted = false;
                if (!widthsEqual)
                {
                    WordScoreExplanation.Note("表の列幅がすべて同じになっていません。");
                    noted = true;
                }
                // 操作記録は見た目（列幅）が揃っているときだけ出す（4-6型）
                if (widthsEqual && !logOk)
                {
                    WordScoreExplanation.Note("列幅を揃えた操作の記録がありません。");
                    noted = true;
                }
                if (noted)
                    return false;
                return Miss("表の列幅が指定どおりになっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_06_Impl(string filePath)
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

                string tableXml = GetTableXmlNearText(xml, "北海道のうまいもの市場売上");
                if (string.IsNullOrEmpty(tableXml))
                    return Miss("見出し「2.北海道のうまいもの市場売上」の下の表が見つかりません。");

                var firstRowMatch = System.Text.RegularExpressions.Regex.Match(
                    tableXml, @"<w:tr[ >].*?</w:tr>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!firstRowMatch.Success)
                    return Miss("表の1行目が見つかりません。");
                string firstRowXml = firstRowMatch.Value;

                var trPrMatch = System.Text.RegularExpressions.Regex.Match(
                    firstRowXml, @"<w:trPr\b[^>]*>.*?</w:trPr>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!trPrMatch.Success)
                    return Miss("表の最初の行が各ページのタイトル行として表示されるようになっていません。");
                string trPrXml = trPrMatch.Value;

                bool hasTblHeader = System.Text.RegularExpressions.Regex.IsMatch(
                    trPrXml, @"<w:tblHeader\b");

                if (hasTblHeader)
                    return true;
                return Miss("表の最初の行が各ページのタイトル行として表示されるようになっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_07_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                Tables tables = document.Tables;
                try
                {
                    int tableCount = 0;
                    try { tableCount = tables.Count; } catch { }

                    // 「北海道のうまいもの市場売上」直下の表より後にある表 = 文書末付近に追加した6-7用の表
                    int anchorIndex = FindFirstTableIndexAfterText(document, tables, "北海道のうまいもの市場売上");
                    if (anchorIndex < 1 || tableCount <= anchorIndex)
                    {
                        // 末尾への表挿入なし（既存表だけ）
                        return Miss("13行3列の表が作成されていません。");
                    }

                    Table endTable = null;
                    try
                    {
                        endTable = tables[tableCount];
                        bool sizeOk = false;
                        try { sizeOk = endTable.Rows.Count == 13 && endTable.Columns.Count == 3; } catch { }

                        bool headersExact = false;
                        bool hasHeaderSample = false;
                        string c1 = "", c2 = "", c3 = "";
                        if (TryReadRow1ThreeCells(endTable, out c1, out c2, out c3))
                        {
                            hasHeaderSample = true;
                            headersExact = AreTask67HeadersExact(c1, c2, c3);
                        }

                        if (sizeOk && headersExact)
                            return true;

                        bool noted = false;
                        if (!sizeOk)
                        {
                            WordScoreExplanation.Note("13行3列の表が作成されていません。");
                            noted = true;
                        }

                        if (!headersExact)
                        {
                            if (!hasHeaderSample
                                || (string.IsNullOrEmpty(c1) && string.IsNullOrEmpty(c2) && string.IsNullOrEmpty(c3)))
                            {
                                WordScoreExplanation.Note("1行目が入力されていません。");
                            }
                            else if (AreTask67HeaderPermutation(c1, c2, c3))
                            {
                                WordScoreExplanation.Note(
                                    $"1行目の見出しの順番が違います。「{QuoteForStudent(c1)}」「{QuoteForStudent(c2)}」「{QuoteForStudent(c3)}」の順になっています。");
                            }
                            else
                            {
                                WordScoreExplanation.Note(
                                    $"1行目の見出しが「{QuoteForStudent(c1)}」「{QuoteForStudent(c2)}」「{QuoteForStudent(c3)}」になっています。");
                            }
                            noted = true;
                        }

                        if (noted)
                            return false;
                        return Miss("表が指定どおりになっていません。");
                    }
                    finally
                    {
                        if (endTable != null) Marshal.ReleaseComObject(endTable);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(tables);
                }
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>指定テキストを含む位置より後に現れる最初の表の 1-based インデックス。無ければ 0。</summary>
        private static int FindFirstTableIndexAfterText(Document document, Tables tables, string targetText)
        {
            if (document == null || tables == null || string.IsNullOrEmpty(targetText))
                return 0;

            Range searchRange = null;
            Find find = null;
            try
            {
                searchRange = document.Content.Duplicate;
                find = searchRange.Find;
                find.ClearFormatting();
                find.Text = targetText;
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;
                if (!find.Execute())
                    return 0;

                int afterPos = searchRange.End;
                int tableCount = 0;
                try { tableCount = tables.Count; } catch { return 0; }

                for (int i = 1; i <= tableCount; i++)
                {
                    Table table = null;
                    try
                    {
                        table = tables[i];
                        if (table.Range.Start >= afterPos)
                            return i;
                    }
                    catch { }
                    finally
                    {
                        if (table != null) Marshal.ReleaseComObject(table);
                    }
                }
                return 0;
            }
            catch
            {
                return 0;
            }
            finally
            {
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
            }
        }

        private static bool AreTask67HeadersExact(string c1, string c2, string c3)
        {
            return string.Equals(c1, "最高売上", StringComparison.Ordinal)
                && string.Equals(c2, "平均売上", StringComparison.Ordinal)
                && string.Equals(c3, "合計売上", StringComparison.Ordinal);
        }

        /// <summary>3語は正解セットだが順番だけ違う。</summary>
        private static bool AreTask67HeaderPermutation(string c1, string c2, string c3)
        {
            if (AreTask67HeadersExact(c1, c2, c3))
                return false;
            var expected = new[] { "最高売上", "平均売上", "合計売上" };
            var actual = new[] { c1 ?? "", c2 ?? "", c3 ?? "" };
            Array.Sort(expected, StringComparer.Ordinal);
            Array.Sort(actual, StringComparer.Ordinal);
            return string.Equals(expected[0], actual[0], StringComparison.Ordinal)
                && string.Equals(expected[1], actual[1], StringComparison.Ordinal)
                && string.Equals(expected[2], actual[2], StringComparison.Ordinal);
        }

        private static bool TryReadRow1ThreeCells(Table table, out string c1, out string c2, out string c3)
        {
            c1 = c2 = c3 = "";
            if (table == null)
                return false;
            try
            {
                if (table.Rows.Count < 1 || table.Columns.Count < 3)
                    return false;
            }
            catch
            {
                return false;
            }

            c1 = ReadCleanCellText(table, 1, 1);
            c2 = ReadCleanCellText(table, 1, 2);
            c3 = ReadCleanCellText(table, 1, 3);
            return true;
        }

        private static string ReadCleanCellText(Table table, int row, int col)
        {
            Cell cell = null;
            try
            {
                cell = table.Cell(row, col);
                return CleanTableCellText(cell.Range.Text);
            }
            catch
            {
                return "";
            }
            finally
            {
                if (cell != null) Marshal.ReleaseComObject(cell);
            }
        }

        /// <summary>表セルの Range.Text から末尾制御文字を除き Trim。</summary>
        private static string CleanTableCellText(string raw)
        {
            if (string.IsNullOrEmpty(raw))
                return "";
            var sb = new System.Text.StringBuilder(raw.Length);
            foreach (char c in raw)
            {
                if (c == '\r' || c == '\n' || c == '\a' || c == '\v' || c == '\f' || c == '\u0007')
                    continue;
                sb.Append(c);
            }
            return sb.ToString().Trim();
        }

        private static string QuoteForStudent(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "（空）";
            string text = value.Replace("\r", "").Replace("\n", " ");
            const int maxLen = 40;
            if (text.Length <= maxLen)
                return text;
            return text.Substring(0, maxLen) + "…";
        }

        /// <summary>
        /// 指定したテキストを含む段落の直後にある表のXMLブロックを抽出します。
        /// </summary>
        private string GetTableXmlNearText(string fullXml, string targetText)
        {
            var paragraphs = System.Text.RegularExpressions.Regex.Matches(fullXml, @"<w:p\b[^>]*>.*?</w:p>", System.Text.RegularExpressions.RegexOptions.Singleline);

            for (int i = 0; i < paragraphs.Count; i++)
            {
                string pXml = paragraphs[i].Value;
                string pText = System.Text.RegularExpressions.Regex.Replace(pXml, @"<[^>]+>", "");

                if (pText.Contains(targetText))
                {
                    int startIndex = paragraphs[i].Index + paragraphs[i].Length;
                    var tableMatch = System.Text.RegularExpressions.Regex.Match(fullXml.Substring(startIndex), @"<w:tbl\b[^>]*>.*?</w:tbl>", System.Text.RegularExpressions.RegexOptions.Singleline);

                    if (tableMatch.Success && tableMatch.Index < 10000)
                    {
                        return tableMatch.Value;
                    }
                }
            }
            return "";
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
