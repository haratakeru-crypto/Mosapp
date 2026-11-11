using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_9
    {
        public bool CheckTask_1_9_01() => RunCheck(CheckTask_1_9_01_Impl);
        public bool CheckTask_1_9_02() => RunCheck(CheckTask_1_9_02_Impl);
        public bool CheckTask_1_9_03() => RunCheck(CheckTask_1_9_03_Impl);
        public bool CheckTask_1_9_04() => RunCheck(CheckTask_1_9_04_Impl);
        public bool CheckTask_1_9_05() => RunCheck(CheckTask_1_9_05_Impl);
        public bool CheckTask_1_9_06() => RunCheck(CheckTask_1_9_06_Impl);

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

        private bool CheckTask_1_9_01_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                // 完成形: 1.ゴールデンウィーク / 2.こどもの日 / 3.母の日
                if (!TryGetListValueAtPhrase(document, "ゴールデンウィーク", out int gwValue))
                    return Miss("段落番号が「1」から開始されていません。");
                if (gwValue != 1)
                    return Miss($"段落番号が「{gwValue}」から始まっています。");

                if (!TryGetListValueAtPhrase(document, "こどもの日", out int kodomoValue)
                    || !TryGetListValueAtPhrase(document, "母の日", out int hahaValue))
                    return Miss("続く段落番号が「2」「3」になっていません。");

                if (kodomoValue == 2 && hahaValue == 3)
                    return true;

                return Miss($"続く段落番号が「{kodomoValue}」「{hahaValue}」になっています。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private static bool TryGetListValueAtPhrase(Document document, string phrase, out int listValue)
        {
            listValue = 0;
            Range searchRange = null;
            Find find = null;
            Paragraphs paragraphs = null;
            Paragraph firstPara = null;
            Range paraRange = null;
            ListFormat lf = null;
            try
            {
                searchRange = WordFindHelper.DuplicateContent(document);
                find = searchRange.Find;
                WordFindHelper.ConfigureSafeFind(find, phrase);
                find.Execute();
                if (!find.Found)
                    return false;

                paragraphs = searchRange.Paragraphs;
                firstPara = paragraphs[1];
                paraRange = firstPara.Range;
                lf = paraRange.ListFormat;
                listValue = lf.ListValue;
                return true;
            }
            catch
            {
                listValue = 0;
                return false;
            }
            finally
            {
                if (lf != null) Marshal.ReleaseComObject(lf);
                if (paraRange != null) Marshal.ReleaseComObject(paraRange);
                if (firstPara != null) Marshal.ReleaseComObject(firstPara);
                if (paragraphs != null) Marshal.ReleaseComObject(paragraphs);
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
            }
        }

        private bool CheckTask_1_9_02_Impl(string filePath)
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

                var documentPartMatch = Regex.Match(xml, @"<pkg:part pkg:name=""/word/document\.xml""[^>]*>.*?</pkg:part>", RegexOptions.Singleline);
                string bodyXml = documentPartMatch.Success ? documentPartMatch.Value : xml;

                var paragraphs = Regex.Matches(bodyXml, @"<w:p\b[^>]*>.*?</w:p>", RegexOptions.Singleline);

                int headingIndex = -1;
                int headingLength = 0;
                int nextHeadingIndex = -1;

                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var p = paragraphs[i];
                    string pText = Regex.Replace(p.Value, @"<[^>]+>", "");

                    if (headingIndex == -1 && pText.Contains("月別催事内容"))
                    {
                        headingIndex = p.Index;
                        headingLength = p.Length;
                    }
                    else if (headingIndex != -1 && nextHeadingIndex == -1 && pText.Contains("小豆島"))
                    {
                        nextHeadingIndex = p.Index;
                        break;
                    }
                }

                if (headingIndex == -1)
                    return Miss("見出し「2. 月別催事内容」が見つかりません。");

                string rangeXml;
                if (nextHeadingIndex != -1 && nextHeadingIndex > headingIndex + headingLength)
                    rangeXml = bodyXml.Substring(headingIndex + headingLength, nextHeadingIndex - (headingIndex + headingLength));
                else
                    rangeXml = bodyXml.Substring(headingIndex + headingLength);

                if (rangeXml.Contains("<w:tbl"))
                    return Miss("表がコンマ区切りの文字列に変更されていません。");

                string rangeText = Regex.Replace(rangeXml, @"<[^>]+>", "");
                int commaCount = 0;
                foreach (char c in rangeText)
                {
                    if (c == ',' || c == '，') commaCount++;
                }

                if (commaCount >= 5)
                    return true;

                // 表の解除ダイアログで［タブ］［段落記号］を選んだ場合
                int tabCount = Regex.Matches(rangeXml, @"<w:tab\b").Count;
                if (tabCount >= 5)
                    return Miss("タブ区切りになっています。");

                int paragraphCount = Regex.Matches(rangeXml, @"<w:p\b").Count;
                // 初期表は2列×7行程度。段落記号区切りだとセルごとに段落が増える
                if (paragraphCount >= 10 && tabCount < 5)
                    return Miss("段落記号区切りになっています。");

                return Miss("コンマ区切りの文字列になっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_9_03_Impl(string filePath)
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

                var documentPartMatch = Regex.Match(xml, @"<pkg:part pkg:name=""/word/document\.xml""[^>]*>.*?</pkg:part>", RegexOptions.Singleline);
                string bodyXml = documentPartMatch.Success ? documentPartMatch.Value : xml;

                var paragraphs = Regex.Matches(bodyXml, @"<w:p\b[^>]*>.*?</w:p>", RegexOptions.Singleline);

                int headingIndex = -1;
                int headingLength = 0;

                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var p = paragraphs[i];
                    string pText = Regex.Replace(p.Value, @"<[^>]+>", "");

                    if (pText.Contains("小豆島のうまいもの市場売上"))
                    {
                        headingIndex = p.Index;
                        headingLength = p.Length;
                        break;
                    }
                }

                if (headingIndex == -1)
                    return Miss("見出し「3. 小豆島のうまいもの市場売上」が見つかりません。");

                string afterHeadingXml = bodyXml.Substring(headingIndex + headingLength);
                var tblMatch = Regex.Match(afterHeadingXml, @"<w:tbl\b[^>]*>.*?</w:tbl>", RegexOptions.Singleline);
                if (!tblMatch.Success)
                    return Miss("見出し「3. 小豆島のうまいもの市場売上」の下の表が見つかりません。");
                string tableXml = tblMatch.Value;

                // 4mm ≒ 227 dxa (誤差許容範囲 220 〜 235)
                int? observedRight = null;
                var marMatches = Regex.Matches(tableXml, @"<(w:tblCellMar|w:tcMar)\b[^>]*>.*?</\1>", RegexOptions.Singleline);
                foreach (Match mar in marMatches)
                {
                    var rightMatch = Regex.Match(mar.Value, @"<w:right\b[^>]*w:w=""(\d+)""");
                    if (rightMatch.Success && int.TryParse(rightMatch.Groups[1].Value, out int width))
                    {
                        observedRight = width;
                        if (width >= 220 && width <= 235)
                            return true;
                    }
                }

                if (observedRight.HasValue)
                    return Miss($"セルの右の余白が「{FormatDxaAsMm(observedRight.Value)}」になっています。");
                return Miss("セルの右の余白が「4mm」になっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_9_04_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            Range searchRange = null; Find find = null;
            Tables tables = null; Table table = null;
            Rows rows = null; Row headerRow = null;
            Cells cells = null; Cell hCell = null;
            Range hCellRange = null; Rows tblRows = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                searchRange = WordFindHelper.DuplicateContent(document);
                find = searchRange.Find;
                WordFindHelper.ConfigureSafeFind(find, "合計");
                find.Execute();

                if (!find.Found)
                    return Miss("表が「合計」の降順に並べ替えられていません。");

                tables = searchRange.Tables;
                if (tables == null || tables.Count < 1)
                    return Miss("表が「合計」の降順に並べ替えられていません。");

                table = tables[1];
                rows = table.Rows;
                headerRow = rows[1];
                cells = headerRow.Cells;
                int numCells = cells.Count;
                int colIndex = -1;

                for (int c = 1; c <= numCells; c++)
                {
                    hCell = cells[c];
                    hCellRange = hCell.Range;
                    try
                    {
                        string cellText = hCellRange?.Text?.Trim() ?? "";
                        if (cellText.Contains("合計")) { colIndex = c; break; }
                    }
                    finally
                    {
                        if (hCellRange != null) Marshal.ReleaseComObject(hCellRange);
                        if (hCell != null) Marshal.ReleaseComObject(hCell);
                        hCellRange = null;
                        hCell = null;
                    }
                }

                if (colIndex < 0)
                    return Miss("表が「合計」の降順に並べ替えられていません。");

                tblRows = table.Rows;
                int rowCount = tblRows.Count;
                var values = new System.Collections.Generic.List<float>();

                for (int r = 2; r <= rowCount; r++)
                {
                    Cell cell = table.Cell(r, colIndex);
                    Range cellRange = cell.Range;
                    try
                    {
                        string cellText = cellRange?.Text?.Trim().TrimEnd('\r', '\a') ?? "";
                        if (string.IsNullOrEmpty(cellText)) continue;
                        if (float.TryParse(cellText.Replace(",", ""), out float val))
                            values.Add(val);
                    }
                    finally
                    {
                        if (cellRange != null) Marshal.ReleaseComObject(cellRange);
                        if (cell != null) Marshal.ReleaseComObject(cell);
                    }
                }

                if (values.Count < 2)
                    return Miss("表が「合計」の降順に並べ替えられていません。");

                bool descending = true;
                bool ascending = true;
                for (int i = 1; i < values.Count; i++)
                {
                    if (values[i] > values[i - 1]) descending = false;
                    if (values[i] < values[i - 1]) ascending = false;
                }

                if (descending)
                    return true;
                if (ascending)
                    return Miss("「合計」の「昇順」になっています。");

                return Miss("表が「合計」の降順に並べ替えられていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally
            {
                if (tblRows != null) Marshal.ReleaseComObject(tblRows);
                if (cells != null) Marshal.ReleaseComObject(cells);
                if (headerRow != null) Marshal.ReleaseComObject(headerRow);
                if (rows != null) Marshal.ReleaseComObject(rows);
                if (table != null) Marshal.ReleaseComObject(table);
                if (tables != null) Marshal.ReleaseComObject(tables);
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_9_05_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            Revisions revisions = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                bool noRevisions = false;
                try
                {
                    revisions = document.Revisions;
                    noRevisions = revisions.Count == 0;
                }
                catch
                {
                    noRevisions = true;
                }

                // 後続の変更履歴ロックで TrackRevisions が戻るため、操作ログと履歴0件で判定する。
                bool stopTrackingOk = LogReader.HasTaskEvidence(9, 5, "AcceptAllChangesInDocAndStopTracking");
                bool acceptOnly = LogReader.HasTaskEvidence(9, 5, "ReviewAcceptAllChangesInDocument");
                bool trackOffLogged = LogReader.HasTaskEvidence(9, 5, "TrackRevisionsOff");
                // 一発正答、または 反映＋変更の記録OFF の二段正答
                bool logOk = stopTrackingOk || (acceptOnly && trackOffLogged);
                if (logOk && noRevisions)
                    return true;

                bool noted = false;
                if (!noRevisions)
                {
                    WordScoreExplanation.Note("変更履歴が残っています。");
                    noted = true;
                }
                if (!logOk)
                {
                    if (noRevisions && acceptOnly && !trackOffLogged && !stopTrackingOk)
                        WordScoreExplanation.Note("変更の記録が停止されていません。");
                    else
                        WordScoreExplanation.Note("すべての変更履歴を反映し、変更の記録を停止した記録がありません。");
                    noted = true;
                }
                if (noted)
                    return false;
                return Miss("変更履歴の承認と記録の停止が完了していません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally
            {
                if (revisions != null) Marshal.ReleaseComObject(revisions);
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_9_06_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                if (document.ProtectionType != WdProtectionType.wdAllowOnlyRevisions)
                    return Miss("変更履歴を編集できないようにロックされていません。");

                // 空パスワードは Unprotect("456") も成功するため、先に空解除を試す
                try
                {
                    document.Unprotect(string.Empty);
                    try
                    {
                        document.Protect(WdProtectionType.wdAllowOnlyRevisions, NoReset: true, Password: string.Empty);
                    }
                    catch { /* ignore */ }
                    return Miss("パスワードが「456」になっていません。");
                }
                catch (COMException)
                {
                    // 何らかのパスワードあり → 続けて 456 を検証
                }

                bool passwordOk = false;
                try
                {
                    document.Unprotect("456");
                    passwordOk = true;
                    document.Protect(WdProtectionType.wdAllowOnlyRevisions, NoReset: true, Password: "456");
                }
                catch (COMException)
                {
                    passwordOk = false;
                }

                if (passwordOk)
                    return true;

                return Miss("パスワードが「456」になっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>twip(dxa) を学生向けの mm 表示に変換（1mm ≒ 56.7 dxa）。</summary>
        private static string FormatDxaAsMm(int dxa)
        {
            double mm = dxa / 56.7;
            string text = mm.ToString("0.#");
            return text + "mm";
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
