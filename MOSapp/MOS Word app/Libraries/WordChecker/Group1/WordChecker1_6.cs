using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_6
    {
        public bool CheckTask_1_6_01() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_01(filePath); } catch { return false; } }
        public bool CheckTask_1_6_02() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_02(filePath); } catch { return false; } }
        public bool CheckTask_1_6_03() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_03(filePath); } catch { return false; } }
        public bool CheckTask_1_6_04() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_04(filePath); } catch { return false; } }
        public bool CheckTask_1_6_05() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_05(filePath); } catch { return false; } }
        public bool CheckTask_1_6_06() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_06(filePath); } catch { return false; } }
        public bool CheckTask_1_6_07() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_07(filePath); } catch { return false; } }

        private bool CheckTask_1_6_01(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // VSTOログから手順をチェック: TableConvertTextToTableが実行されたか
                string logFilePath = LogReader.GetLogFilePath();
                bool logFileExists = System.IO.File.Exists(logFilePath);
                bool tableConvertExecuted = LogReader.HasCommandExecuted("TableConvertTextToTable");
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_6_01] Log file exists: {logFileExists}");
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_6_01] TableConvertTextToTable executed: {tableConvertExecuted}");

                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "（2）月別催事内容"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                // 7行2列の表が存在するかチェック
                Tables tables = document.Tables;
                bool fileStateCheck = false;
                foreach (Table table in tables) { if (table.Rows.Count == 7 && table.Columns.Count == 2) { fileStateCheck = true; break; } Marshal.ReleaseComObject(table); }
                Marshal.ReleaseComObject(tables); Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);

                // Phase1: ログ優先（変換ログ + 文書状態の両方が必要）
                bool result = logFileExists && tableConvertExecuted && fileStateCheck;
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_6_01] Result (log + file state): {result}");
                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_02(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「（2）月別催事内容」付近の表を5行目から分割した結果のみを対象にし、さらにVSTOログ（SplitTable）がある場合のみ正解とする
                Range headingRange = null;
                bool hasTargetHeading = false;
                string[] targets = new[] { "（2）月別催事内容", "2. 月別催事内容", "２. 月別催事内容", "月別催事内容" };
                foreach (string t in targets)
                {
                    Range searchRange = document.Content;
                    Find find = searchRange.Find;
                    find.ClearFormatting();
                    find.Text = t;
                    find.Execute();
                    hasTargetHeading = find.Found;
                    if (hasTargetHeading)
                    {
                        headingRange = searchRange.Duplicate;
                        Marshal.ReleaseComObject(find);
                        Marshal.ReleaseComObject(searchRange);
                        break;
                    }
                    Marshal.ReleaseComObject(find);
                    Marshal.ReleaseComObject(searchRange);
                }
                if (!hasTargetHeading || headingRange == null) return false;

                int headingEnd = headingRange.End;
                Tables tables = document.Tables;
                Table table1 = null;
                Table table2 = null;
                try
                {
                    for (int i = 1; i <= tables.Count; i++)
                    {
                        Table t = tables[i];
                        int start = t.Range.Start;
                        if (start <= headingEnd)
                        {
                            Marshal.ReleaseComObject(t);
                            continue;
                        }
                        if (table1 == null)
                        {
                            table1 = t; // 後で解放
                        }
                        else
                        {
                            table2 = t;
                            Marshal.ReleaseComObject(t);
                            break;
                        }
                    }

                    bool stateOk = false;
                    if (table1 != null && table2 != null)
                    {
                        // 「5行目から分割」した結果、先頭表の行数が4行になっていることをチェック
                        try
                        {
                            stateOk = table1.Rows.Count == 4;
                        }
                        catch { stateOk = false; }
                    }

                    bool logOk = LogReader.HasCommandExecuted("SplitTable");
                    return stateOk && logOk;
                }
                finally
                {
                    if (table1 != null) Marshal.ReleaseComObject(table1);
                    if (table2 != null) Marshal.ReleaseComObject(table2);
                    Marshal.ReleaseComObject(tables);
                    Marshal.ReleaseComObject(headingRange);
                }
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_03(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 1つ目の表の1行2列目を3列に分割 → 1行目のセル数が4になる（1列目+分割3列）
                Tables tables = document.Tables;
                bool result = false;
                if (tables.Count >= 1)
                {
                    Table t = tables[1];
                    if (t.Rows.Count >= 1 && t.Rows[1].Cells.Count >= 4) result = true;
                    Marshal.ReleaseComObject(t);
                }
                Marshal.ReleaseComObject(tables);
                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_04(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「2.」は段落番号で文字ではないため、「北海道のうまいもの市場売上」で検索
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "北海道のうまいもの市場売上"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                // 表の1行目に塗りつぶし・輪郭（オレンジ系）が設定されているか
                Table table = searchRange.Tables[1];
                Row firstRow = table.Rows[1];
                Cell firstCell = firstRow.Cells[1];
                Range cellRange = firstCell.Range;
                Font font = cellRange.Font;
                bool hasEffect = (int)font.Color != (int)WdColor.wdColorAutomatic || (int)firstCell.Range.ParagraphFormat.Shading.BackgroundPatternColor != (int)WdColor.wdColorAutomatic;
                Marshal.ReleaseComObject(font); Marshal.ReleaseComObject(cellRange); Marshal.ReleaseComObject(firstCell); Marshal.ReleaseComObject(firstRow); Marshal.ReleaseComObject(table);
                Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                return hasEffect;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_05(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 対象表（「北海道のうまいもの市場売上」を含む）の列幅が等しいか
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "北海道のうまいもの市場売上"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Table table = searchRange.Tables[1];
                Columns cols = table.Columns;
                int colCount = cols.Count;
                float firstWidth = colCount >= 1 ? cols[1].Width : 0;
                const float tolerance = 2f; // 2pt の誤差を許容
                bool equal = colCount >= 2;
                for (int c = 2; c <= colCount && equal; c++)
                {
                    if (Math.Abs(cols[c].Width - firstWidth) > tolerance) equal = false;
                }
                Marshal.ReleaseComObject(cols); Marshal.ReleaseComObject(table); Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                // Phase1: ログ優先（操作ログ + 文書状態の両方が必要）
                return LogReader.HasCommandExecuted("TableColumnsDistribute") && equal;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_06(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "2.北海道のうまいもの市場売上"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                // タイトル行の繰り返しが設定されているかチェック
                Tables tables = document.Tables;
                bool result = false;
                foreach (Table table in tables)
                {
                    if (table.Rows[1].HeadingFormat != 0) { result = true; break; }
                    Marshal.ReleaseComObject(table);
                }
                Marshal.ReleaseComObject(tables); Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                // Phase1: ログ優先（操作ログ + 文書状態の両方が必要）
                return LogReader.HasCommandExecuted("TableRepeatHeaderRows") && result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_07(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 13行3列の表が存在し、1行目に「最高売上」「平均売上」「合計売上」があるかチェック
                Tables tables = document.Tables;
                bool result = false;
                foreach (Table table in tables)
                {
                    if (table.Rows.Count == 13 && table.Columns.Count == 3)
                    {
                        string row1Text = table.Rows[1].Range.Text;
                        if (row1Text.Contains("最高売上") && row1Text.Contains("平均売上") && row1Text.Contains("合計売上")) { result = true; break; }
                    }
                    Marshal.ReleaseComObject(table);
                }
                Marshal.ReleaseComObject(tables); return result;
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

