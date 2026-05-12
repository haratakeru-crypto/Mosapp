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
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「4.」は段落番号で文字ではないため、「ゴールデンウィーク」で検索
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "ゴールデンウィーク"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Range foundRange = searchRange;
                ListFormat lf = null;
                bool result = false;
                try
                {
                    lf = foundRange.Paragraphs[1].Range.ListFormat;
                    result = lf.ListValue == 1;
                }
                finally { if (lf != null) Marshal.ReleaseComObject(lf); Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); }
                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_9_02(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「月別催事内容」の直後がコンマ区切りテキスト（表を文字列に変換した結果）か
                string[] targets = new[] { "（2）月別催事内容", "2. 月別催事内容", "２. 月別催事内容", "月別催事内容" };
                Range foundRange = null;
                foreach (string t in targets)
                {
                    Range searchRange = document.Content;
                    Find find = searchRange.Find;
                    find.ClearFormatting();
                    find.Text = t;
                    find.Execute();
                    if (find.Found)
                    {
                        foundRange = searchRange.Duplicate;
                        Marshal.ReleaseComObject(find);
                        Marshal.ReleaseComObject(searchRange);
                        break;
                    }
                    Marshal.ReleaseComObject(find);
                    Marshal.ReleaseComObject(searchRange);
                }
                if (foundRange == null) return false;

                Range afterRange = foundRange.Duplicate;
                afterRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                // 直後の段落～次段落あたりを取得（表→文字列変換後は段落構造が変わるため広めに取得）
                afterRange.MoveEnd(WdUnits.wdParagraph, 2);
                string afterText = afterRange.Text ?? "";
                Marshal.ReleaseComObject(afterRange);
                Marshal.ReleaseComObject(foundRange);
                // コンマが2つ以上あり、行らしき区切り（改行やタブ）がある
                int commaCount = 0; foreach (char c in afterText) if (c == ',') commaCount++;
                return commaCount >= 2 && (afterText.Contains("\r") || afterText.Contains("\n") || afterText.Contains(","));
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_9_03(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「合計」を含む表のセル右余白が約4mm（約11.33pt）か
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "合計"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Table table = searchRange.Tables[1];
                int row = table.Rows.Count >= 2 ? 2 : 1;
                Cell cell = table.Cell(row, 1);
                float rightIndent = cell.Range.ParagraphFormat.RightIndent;
                Marshal.ReleaseComObject(cell); Marshal.ReleaseComObject(table); Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                const float fourMmPt = 11.33f; const float tolerance = 2f;
                return Math.Abs(rightIndent - fourMmPt) <= tolerance;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_9_04(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「合計」列を探し、その列の値が降順か
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "合計"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Table table = searchRange.Tables[1];
                int colIndex = 1; Row headerRow = table.Rows[1];
                int numCells = headerRow.Cells.Count;
                for (int c = 1; c <= numCells; c++)
                {
                    Cell hCell = headerRow.Cells[c];
                    try
                    {
                        string cellText = hCell.Range?.Text?.Trim() ?? "";
                        if (cellText.Contains("合計")) { colIndex = c; break; }
                    }
                    finally { Marshal.ReleaseComObject(hCell); }
                }
                Marshal.ReleaseComObject(headerRow);
                float prev = float.MaxValue;
                bool descending = table.Rows.Count >= 2;
                for (int r = 2; r <= table.Rows.Count && descending; r++)
                {
                    Cell cell = table.Cell(r, colIndex);
                    string cellText = cell.Range?.Text?.Trim().TrimEnd('\r', '\a') ?? "";
                    Marshal.ReleaseComObject(cell);
                    if (string.IsNullOrEmpty(cellText)) continue;
                    if (float.TryParse(cellText.Replace(",", ""), out float val))
                    {
                        if (val > prev) descending = false;
                        prev = val;
                    }
                }
                Marshal.ReleaseComObject(table); Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                return descending;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_9_05(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                bool trackOff = !document.TrackRevisions;
                bool noRevisions = true;
                try { noRevisions = document.Revisions.Count == 0; } catch { }
                bool stateOk = trackOff && noRevisions;
                // Phase1: ログ優先（全承認ログ + 文書状態の両方が必要）
                return LogReader.HasCommandExecuted("ReviewAcceptAllChangesInDocument") && stateOk;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_9_06(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 変更履歴のロック: 編集制限で「変更履歴のみ」が有効か
                return document.ProtectionType == WdProtectionType.wdAllowOnlyRevisions;
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

