using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_1
    {
        public bool CheckTask_1_1_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_1_01(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        public bool CheckTask_1_1_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_1_02(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        /// <summary>
        /// Project9のタスク9-3をチェックする
        /// シート[下半期売上]の「7月」から「12月」の列に、アイコンセット「3つの矢印（色分け）」を設定
        /// </summary>
        public bool CheckTask_1_1_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_1_03(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        public bool CheckTask_1_1_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_1_04(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        public bool CheckTask_1_1_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_1_05(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        public bool CheckTask_1_1_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_1_06(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        public bool CheckTask_1_1_07()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_1_07(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }



        // ==========================================
        // タスク1-1: 印刷の向き（横）
        // 条件：「売上一覧」のみ横向き(xlLandscape)。他は縦向き(xlPortrait)であること。
        // ==========================================
        private bool CheckTask_1_1_01(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;

            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application { Visible = false };
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                bool targetSheetOk = false;
                var landscapeOthers = new List<string>();

                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // シートごとの印刷の向きを取得
                    XlPageOrientation orientation = ws.PageSetup.Orientation;

                    if (ws.Name == "売上一覧")
                    {
                        // 対象シートは「横向き(xlLandscape)」ならOK
                        if (orientation == XlPageOrientation.xlLandscape)
                        {
                            targetSheetOk = true;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[Task 1-1] Target sheet '{ws.Name}' is not Landscape.");
                        }
                    }
                    else
                    {
                        // 他のシートは「横向き」になっていたらNG（＝縦向きであるべき）
                        if (orientation == XlPageOrientation.xlLandscape)
                        {
                            landscapeOthers.Add(ws.Name);
                            System.Diagnostics.Debug.WriteLine($"[Task 1-1] Other sheet '{ws.Name}' is Landscape (Should be Portrait).");
                        }
                    }
                }

                if (targetSheetOk && landscapeOthers.Count == 0)
                    return true;
                if (!targetSheetOk)
                    ExcelScoreExplanation.Note("シート「売上一覧」の印刷の向きが縦になっています。");
                if (landscapeOthers.Count > 0)
                    ExcelScoreExplanation.Note($"シート「{JoinNames(landscapeOthers)}」の印刷の向きが横になっています。");
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Task 1-1] Error: {ex.Message}");
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (workbook != null)
                {
                    try { Marshal.ReleaseComObject(workbook); } catch { }
                }
            }
        }

        private bool CheckTask_1_1_02(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                // 既に開いているExcelアプリケーションを取得
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[ExcelChecker1_1_02] Excelアプリケーションが見つかりません");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }
                
                // デバッグログ追加
                System.Diagnostics.Debug.WriteLine($"[ExcelChecker1_1_02] Looking for file: {filePath}");
                
                // 既に開いているワークブックを検索
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelChecker1_1_02] Workbook not found (path/name match).");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }

                // 売上一覧シートを検索
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                return IsSalesListPrintAreaAcceptable(worksheet, worksheet.PageSetup.PrintArea);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        /// <summary>
        /// 3-2: 印刷範囲の本体は A4:F131。メモで右方向の長方形に広がった場合は、
        /// 5行目以降に値がないときだけ許可する。4行目の見出しは見ない。
        /// </summary>
        private static bool IsSalesListPrintAreaAcceptable(Worksheet worksheet, string printArea)
        {
            if (worksheet == null)
                return Miss(ExcelScoreExplanation.UnavailableText);
            if (string.IsNullOrWhiteSpace(printArea))
                return Miss("印刷範囲が設定されていません。");

            string normalized = printArea.Replace(" ", "").Replace("$", "").ToUpperInvariant();
            int bang = normalized.LastIndexOf('!');
            if (bang >= 0)
                normalized = normalized.Substring(bang + 1);
            if (normalized.IndexOf(',') >= 0)
                return Miss($"印刷範囲が複数に分かれています（いまの設定: 「{Quote(normalized)}」）。");
            if (!TryParseA1Range(normalized, out int startCol, out int startRow, out int endCol, out int endRow))
                return Miss($"印刷範囲が「{Quote(normalized)}」になっています。");
            if (startCol != 1 || startRow != 4 || endRow != 131 || endCol < 6)
                return Miss($"印刷範囲が「{normalized}」になっています。");
            if (endCol == 6)
                return true;

            Range extra = null;
            try
            {
                extra = worksheet.Range[worksheet.Cells[5, 7], worksheet.Cells[131, endCol]];
                object raw = extra.Value2;
                if (raw == null)
                    return true;
                if (!(raw is object[,] values))
                {
                    if (string.IsNullOrWhiteSpace(Convert.ToString(raw)))
                        return true;
                    return Miss($"印刷範囲が「{normalized}」になっており、指定より広い部分に値が入っています。");
                }

                int rows = values.GetLength(0);
                int cols = values.GetLength(1);
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        if (!string.IsNullOrWhiteSpace(Convert.ToString(values[r, c])))
                            return Miss($"印刷範囲が「{normalized}」になっており、指定より広い部分に値が入っています。");
                    }
                }
                return true;
            }
            catch
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (extra != null)
                {
                    try { Marshal.ReleaseComObject(extra); } catch { }
                }
            }
        }

        private static bool TryParseA1Range(string range, out int startCol, out int startRow, out int endCol, out int endRow)
        {
            startCol = startRow = endCol = endRow = 0;
            string[] parts = range.Split(':');
            if (parts.Length != 2)
                return false;
            return TryParseA1Cell(parts[0], out startCol, out startRow)
                && TryParseA1Cell(parts[1], out endCol, out endRow)
                && startCol <= endCol
                && startRow <= endRow;
        }

        private static bool TryParseA1Cell(string cell, out int column, out int row)
        {
            column = 0;
            row = 0;
            if (string.IsNullOrEmpty(cell))
                return false;

            int index = 0;
            while (index < cell.Length && cell[index] >= 'A' && cell[index] <= 'Z')
            {
                column = (column * 26) + (cell[index] - 'A' + 1);
                index++;
            }
            if (column <= 0 || index >= cell.Length)
                return false;
            return int.TryParse(cell.Substring(index), out row) && row > 0;
        }

        private bool CheckTask_1_1_03(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                // 既に開いているExcelアプリケーションを取得
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[ExcelChecker1_1_03] Excelアプリケーションが見つかりません");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }
                
                // 既に開いているワークブックを検索
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelChecker1_1_03] Workbook not found (path/name match).");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }

                // 売上一覧シートを検索
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);
                
                // 印刷タイトル（タイトル行）の設定をチェック
                string titleRows = worksheet.PageSetup.PrintTitleRows;
                
                // 期待される設定: $2:$4 または 2:4
                if (string.IsNullOrEmpty(titleRows))
                    return Miss("タイトル行が設定されていません。");

                string normalizedTitleRows = titleRows.Replace("$", "").Replace(" ", "").ToUpper();
                if (normalizedTitleRows == "2:4")
                    return true;
                return Miss($"タイトル行が「{normalizedTitleRows}」になっています。");

            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_1_04(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                // 既に開いているExcelアプリケーションを取得
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[ExcelChecker1_1_04] Excelアプリケーションが見つかりません");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }
                
                // 既に開いているワークブックを検索
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelChecker1_1_04] Workbook not found (path/name match).");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }

                // 販売実績シートを検索
                worksheet = FindWorksheet(workbook, "販売実績");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);
                
                // 余白設定をチェック（「広い」の設定値）
                // 広い余白: 上下1インチ(72ポイント)、左右1インチ(72ポイント)
                double topMargin = worksheet.PageSetup.TopMargin;
                double bottomMargin = worksheet.PageSetup.BottomMargin;
                double leftMargin = worksheet.PageSetup.LeftMargin;
                double rightMargin = worksheet.PageSetup.RightMargin;
                
                // 72ポイント（1インチ）の許容範囲をチェック
                const double expectedMargin = 72.0;
                const double tolerance = 1.0;

                var wrongSides = new List<string>();
                if (Math.Abs(topMargin - expectedMargin) > tolerance)
                    wrongSides.Add("上");
                if (Math.Abs(bottomMargin - expectedMargin) > tolerance)
                    wrongSides.Add("下");
                if (Math.Abs(leftMargin - expectedMargin) > tolerance)
                    wrongSides.Add("左");
                if (Math.Abs(rightMargin - expectedMargin) > tolerance)
                    wrongSides.Add("右");

                if (wrongSides.Count == 0)
                    return true;
                if (wrongSides.Count == 4)
                    return Miss("シート「販売実績」の余白が「広い」になっていません。");
                return Miss($"シート「販売実績」の余白（{JoinNames(wrongSides)}）が「広い」になっていません。");
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_1_05(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                // 既に開いているExcelアプリケーションを取得
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[ExcelChecker1_1_05] Excelアプリケーションが見つかりません");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }
                
                // 既に開いているワークブックを検索
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelChecker1_1_05] Workbook not found (path/name match).");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }

                worksheet = FindWorksheet(workbook, "販売実績");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);
                
                // 改ページ位置をチェック
                var vPageBreaks = worksheet.VPageBreaks;
                bool hasCorrectPageBreak = false;
                var manualCols = new List<string>();
                foreach (VPageBreak pageBreak in vPageBreaks)
                {
                    if (pageBreak.Type != XlPageBreak.xlPageBreakManual)
                        continue;
                    int col = pageBreak.Location.Column;
                    if (col == 8)
                    {
                        hasCorrectPageBreak = true;
                        break;
                    }
                    string letter = ColumnLetter(col);
                    if (!manualCols.Contains(letter))
                        manualCols.Add(letter);
                }
                if (hasCorrectPageBreak)
                    return true;
                if (manualCols.Count == 0)
                    return Miss("手動の改ページがありません。");
                return Miss($"手動の改ページが{JoinNames(manualCols.Select(c => c + "列").ToList())}にあります。");
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_1_06(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                // 既に開いているExcelアプリケーションを取得
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[ExcelChecker1_1_06] Excelアプリケーションが見つかりません");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }
                
                // 既に開いているワークブックを検索
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelChecker1_1_06] Workbook not found (path/name match).");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }
                // スキルアップ検定結果シートを検索
                worksheet = FindWorksheet(workbook, "スキルアップ検定結果");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);
                
                // A4:K4範囲の「折り返して全体を表示する」設定をチェック
                Range targetRange = worksheet.Range["A4:K4"];
                var missingWrap = new List<string>();

                // 範囲内のすべてのセルで「折り返して全体を表示する」が設定されているかチェック
                foreach (Range cell in targetRange.Cells)
                {
                    if (!Convert.ToBoolean(cell.WrapText))
                    {
                        string address = cell.Address[false, false];
                        if (!string.IsNullOrEmpty(address) && !missingWrap.Contains(address))
                            missingWrap.Add(address);
                    }
                }

                if (missingWrap.Count == 0)
                    return true;
                return Miss($"{JoinNames(missingWrap)}が折り返して全体を表示するになっていません。");
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_1_07(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            Comments comments = null;
            Range targetCell = null;
            Comment targetComment = null;
            try
            {
                // 既に開いているExcelアプリケーションを取得
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[ExcelChecker1_1_07] Excelアプリケーションが見つかりません");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }
                
                // 既に開いているワークブックを検索
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelChecker1_1_07] Workbook not found (path/name match).");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }

                // 売上一覧シートを検索
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);
                
                // G4セルのメモをチェック
                targetCell = worksheet.Range["G4"];
                targetComment = targetCell.Comment;
                comments = worksheet.Comments;
                var otherAddresses = CollectCommentAddresses(comments, "G4");

                bool g4Ok = targetComment != null;
                if (!g4Ok)
                {
                    ExcelScoreExplanation.Note("指定のセル（売上一覧のG4）にメモがありません。");
                }
                else
                {
                    // メモの内容が「最新の商品情報」かチェック
                    string commentText = targetComment.Text() ?? string.Empty;
                    string normalizedText = commentText.Replace("\r", "").Replace("\n", "");
                    int colonIndex = normalizedText.IndexOf(':');
                    if (colonIndex >= 0 && colonIndex < normalizedText.Length - 1)
                        normalizedText = normalizedText.Substring(colonIndex + 1);

                    if (!string.Equals(normalizedText, "最新の商品情報", StringComparison.Ordinal))
                    {
                        ExcelScoreExplanation.Note($"メモの文章が「{Quote(normalizedText)}」になっています。");
                        g4Ok = false;
                    }
                }

                bool othersOk = otherAddresses.Count == 0;
                if (!othersOk)
                    ExcelScoreExplanation.Note($"指定以外のセル（{JoinNames(otherAddresses.Select(a => "売上一覧の" + a).ToList())}）にもメモがあります。");

                return g4Ok && othersOk;
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (targetComment != null)
                {
                    Marshal.ReleaseComObject(targetComment);
                }
                if (targetCell != null)
                {
                    Marshal.ReleaseComObject(targetCell);
                }
                if (comments != null)
                {
                    Marshal.ReleaseComObject(comments);
                }
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private static bool Miss(string reason)
        {
            ExcelScoreExplanation.Note(reason);
            return false;
        }

        private static string Quote(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "（空）";
            string text = value.Replace("\r", "").Replace("\n", " ");
            const int maxLen = 40;
            if (text.Length <= maxLen)
                return text;
            return text.Substring(0, maxLen) + "…";
        }

        private static string JoinNames(IList<string> names)
        {
            if (names == null || names.Count == 0)
                return "";
            const int maxItems = 5;
            if (names.Count <= maxItems)
                return string.Join("、", names);
            return string.Join("、", names.Take(maxItems)) + "ほか";
        }

        private static string ColumnLetter(int column)
        {
            if (column <= 0)
                return "?";
            var sb = new StringBuilder();
            int n = column;
            while (n > 0)
            {
                n--;
                sb.Insert(0, (char)('A' + (n % 26)));
                n /= 26;
            }
            return sb.ToString();
        }

        private static List<string> CollectCommentAddresses(Comments comments, string excludeAddress)
        {
            var addresses = new List<string>();
            if (comments == null)
                return addresses;

            int count;
            try { count = comments.Count; }
            catch { return addresses; }

            for (int i = 1; i <= count; i++)
            {
                Comment comment = null;
                Range parentRange = null;
                try
                {
                    comment = comments.Item(i);
                    parentRange = comment.Parent as Range;
                    string address = parentRange?.Address[false, false] ?? string.Empty;
                    if (string.IsNullOrEmpty(address))
                        continue;
                    if (!string.IsNullOrEmpty(excludeAddress)
                        && string.Equals(address, excludeAddress, StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (!addresses.Contains(address))
                        addresses.Add(address);
                }
                catch { }
                finally
                {
                    if (parentRange != null)
                    {
                        try { Marshal.ReleaseComObject(parentRange); } catch { }
                    }
                    if (comment != null)
                    {
                        try { Marshal.ReleaseComObject(comment); } catch { }
                    }
                }
            }

            return addresses;
        }

        private string GetCurrentExcelFilePath()
        {
            Application excelApp = null;
            try
            {
                // 実行中のExcelアプリケーションを取得
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                // アクティブなワークブックのパスを取得
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
                
                return null;
            }
            catch (COMException)
            {
                // Excelが起動していない場合
                return null;
            }
            finally
            {
                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);
            }
        }
        
        private Worksheet FindWorksheet(Workbook workbook, string sheetName)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }
            return null;
        }

        private Workbook GetWorkbook(Application excelApp, string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            foreach (Workbook wb in excelApp.Workbooks)
            {
                if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                    wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                {
                    return wb;
                }
            }
            return null;
        }

    }
}