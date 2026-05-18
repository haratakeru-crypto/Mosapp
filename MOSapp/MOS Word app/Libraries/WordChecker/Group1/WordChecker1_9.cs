using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_9
    {
        public bool CheckTask_1_9_01() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_9_01(filePath); } catch { return false; } }
        public bool CheckTask_1_9_02() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_9_02(filePath); } catch { return false; } }
        public bool CheckTask_1_9_03() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_9_03(filePath); } catch { return false; } }
        public bool CheckTask_1_9_04() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_9_04(filePath); } catch { return false; } }
        public bool CheckTask_1_9_05() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_9_05(filePath); } catch { return false; } }
        public bool CheckTask_1_9_06() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_9_06(filePath); } catch { return false; } }

        private bool CheckTask_1_9_01(string filePath)
        {
            Application wordApp = null; Document document = null;
            Range searchRange = null; Find find = null;
            Paragraphs paragraphs = null; Paragraph firstPara = null;
            Range paraRange = null; ListFormat lf = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                searchRange = document.Content;
                find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "ゴールデンウィーク";
                find.Execute();
                
                if (!find.Found) return false;

                paragraphs = searchRange.Paragraphs;
                firstPara = paragraphs[1];
                paraRange = firstPara.Range;
                lf = paraRange.ListFormat;

                return lf.ListValue == 1;
            }
            catch { return false; }
            finally
            {
                if (lf != null) Marshal.ReleaseComObject(lf);
                if (paraRange != null) Marshal.ReleaseComObject(paraRange);
                if (firstPara != null) Marshal.ReleaseComObject(firstPara);
                if (paragraphs != null) Marshal.ReleaseComObject(paragraphs);
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_9_02(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 1. 本文 (/word/document.xml) の抽出
                var documentPartMatch = System.Text.RegularExpressions.Regex.Match(xml, @"<pkg:part pkg:name=""/word/document.xml""[^>]*>.*?</pkg:part>", System.Text.RegularExpressions.RegexOptions.Singleline);
                string bodyXml = documentPartMatch.Success ? documentPartMatch.Value : xml;

                var paragraphs = System.Text.RegularExpressions.Regex.Matches(bodyXml, @"<w:p\b[^>]*>.*?</w:p>", System.Text.RegularExpressions.RegexOptions.Singleline);
                
                int headingIndex = -1;
                int headingLength = 0;
                int nextHeadingIndex = -1;

                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var p = paragraphs[i];
                    string pText = System.Text.RegularExpressions.Regex.Replace(p.Value, @"<[^>]+>", "");
                    
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

                if (headingIndex == -1) return false;

                // 2. 「月別催事内容」見出しから「小豆島」見出しまでの間のXMLを切り出す
                string rangeXml = "";
                if (nextHeadingIndex != -1 && nextHeadingIndex > headingIndex + headingLength)
                {
                    rangeXml = bodyXml.Substring(headingIndex + headingLength, nextHeadingIndex - (headingIndex + headingLength));
                }
                else
                {
                    rangeXml = bodyXml.Substring(headingIndex + headingLength);
                }

                // 3. 該当範囲に表タグ <w:tbl> が含まれていないことを確認
                if (rangeXml.Contains("<w:tbl")) return false;

                // 4. カンマが規定数（5個以上）含まれていることを確認
                string rangeText = System.Text.RegularExpressions.Regex.Replace(rangeXml, @"<[^>]+>", "");
                int commaCount = 0;
                foreach (char c in rangeText)
                {
                    if (c == ',' || c == '，') commaCount++;
                }

                return commaCount >= 5;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_9_03(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 1. 本文 (/word/document.xml) の抽出
                var documentPartMatch = System.Text.RegularExpressions.Regex.Match(xml, @"<pkg:part pkg:name=""/word/document.xml""[^>]*>.*?</pkg:part>", System.Text.RegularExpressions.RegexOptions.Singleline);
                string bodyXml = documentPartMatch.Success ? documentPartMatch.Value : xml;

                var paragraphs = System.Text.RegularExpressions.Regex.Matches(bodyXml, @"<w:p\b[^>]*>.*?</w:p>", System.Text.RegularExpressions.RegexOptions.Singleline);
                
                int headingIndex = -1;
                int headingLength = 0;

                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var p = paragraphs[i];
                    string pText = System.Text.RegularExpressions.Regex.Replace(p.Value, @"<[^>]+>", "");
                    
                    if (pText.Contains("小豆島のうまいもの市場売上"))
                    {
                        headingIndex = p.Index;
                        headingLength = p.Length;
                        break;
                    }
                }

                if (headingIndex == -1) return false;

                // 2. 見出しの直後に出現する表（<w:tbl>）のXMLブロックを抽出
                string afterHeadingXml = bodyXml.Substring(headingIndex + headingLength);
                var tblMatch = System.Text.RegularExpressions.Regex.Match(afterHeadingXml, @"<w:tbl\b[^>]*>.*?</w:tbl>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!tblMatch.Success) return false;
                string tableXml = tblMatch.Value;

                // 3. 表全体のセル余白（tblCellMar）または個別セルの余白（tcMar）から「右余白」の設定値をチェック
                // 4mm ≒ 227 dxa (誤差許容範囲 220 〜 235)
                var marMatches = System.Text.RegularExpressions.Regex.Matches(tableXml, @"<(w:tblCellMar|w:tcMar)\b[^>]*>.*?</\1>", System.Text.RegularExpressions.RegexOptions.Singleline);
                foreach (System.Text.RegularExpressions.Match mar in marMatches)
                {
                    var rightMatch = System.Text.RegularExpressions.Regex.Match(mar.Value, @"<w:right\b[^>]*w:w=""(\d+)""");
                    if (rightMatch.Success && int.TryParse(rightMatch.Groups[1].Value, out int width))
                    {
                        if (width >= 220 && width <= 235)
                        {
                            return true;
                        }
                    }
                }

                return false;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_9_04(string filePath)
        {
            Application wordApp = null; Document document = null;
            Range searchRange = null; Find find = null;
            Tables tables = null; Table table = null;
            Rows rows = null; Row headerRow = null;
            Cells cells = null; Cell hCell = null;
            Range hCellRange = null; Rows tblRows = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                searchRange = document.Content;
                find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "合計";
                find.Execute();
                
                if (!find.Found) return false;

                tables = searchRange.Tables;
                table = tables[1];
                rows = table.Rows;
                headerRow = rows[1];
                cells = headerRow.Cells;
                int numCells = cells.Count;
                int colIndex = 1;

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
                    }
                }

                tblRows = table.Rows;
                int rowCount = tblRows.Count;
                float prev = float.MaxValue;
                bool descending = rowCount >= 2;

                for (int r = 2; r <= rowCount && descending; r++)
                {
                    Cell cell = table.Cell(r, colIndex);
                    Range cellRange = cell.Range;
                    try
                    {
                        string cellText = cellRange?.Text?.Trim().TrimEnd('\r', '\a') ?? "";
                        if (string.IsNullOrEmpty(cellText)) continue;
                        if (float.TryParse(cellText.Replace(",", ""), out float val))
                        {
                            if (val > prev) descending = false;
                            prev = val;
                        }
                    }
                    finally
                    {
                        if (cellRange != null) Marshal.ReleaseComObject(cellRange);
                        if (cell != null) Marshal.ReleaseComObject(cell);
                    }
                }

                return descending;
            }
            catch { return false; }
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

        private bool CheckTask_1_9_05(string filePath)
        {
            Application wordApp = null; Document document = null;
            Revisions revisions = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

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

                // 登録した正しいコマンドログが実行され、かつ未処理の変更履歴が残っていないことを厳格に判定
                // (注: 後続タスクの「変更履歴のロック」を行うと、Wordの仕様により強制的にTrackRevisionsがtrueに戻ってしまい、
                //  状態のみでの追跡が破綻するため、操作ログと履歴0件のANDで完全な厳格性と独立性を担保します)
                return LogReader.HasCommandExecuted("AcceptAllChangesInDocAndStopTracking") && noRevisions;
            }
            catch { return false; }
            finally
            {
                if (revisions != null) Marshal.ReleaseComObject(revisions);
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_9_06(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // 1. まず編集制限タイプが「変更履歴のみ」になっているか
                if (document.ProtectionType != WdProtectionType.wdAllowOnlyRevisions)
                {
                    return false;
                }

                // 2. パスワードが「654」に設定されているかをサイレント検証
                bool passwordOk = false;
                try
                {
                    // パスワード「654」で解除を試みる (画面上のダイアログや警告は一切出ません)
                    document.Unprotect("654");
                    passwordOk = true;

                    // 検証成功したため、即座に同じ「変更履歴のみ」「パスワード: 654」で保護を掛け直す
                    // NoReset: true を指定することで、文書の状態を一切リセットせず完全に維持します
                    document.Protect(WdProtectionType.wdAllowOnlyRevisions, NoReset: true, Password: "654");
                }
                catch (COMException)
                {
                    // パスワードが間違っている、またはパスワードなしでロックされている場合は
                    // 例外が発生するため自動的に不合格とします。文書の保護状態は維持されます。
                    passwordOk = false;
                }

                return passwordOk;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
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

