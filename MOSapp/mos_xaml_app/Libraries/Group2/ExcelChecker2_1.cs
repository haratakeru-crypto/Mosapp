using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace Libraries.Group2
{
    public class ExcelChecker2_1
    {
        public bool CheckTask_2_1_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_1_01_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_1_01_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_1_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_1_02_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_1_02_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_1_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_1_03_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_1_03_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_1_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_1_04_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_1_04_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_1_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_1_05_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_1_05_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_1_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_1_06_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_1_06_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_1_07()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_1_07_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_1_07_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_1_01_Impl(string filePath)
        {
            // ExcelChecker1_1のCheckTask_1_1_01を参考に、他のシートが横向きになっていないかもチェック
            Application excelApp = null;
            Workbook workbook = null;

            try
            {
                // 既に開いているExcelアプリケーションを取得
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = false;
                }
                
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
                
                if (workbook == null) return false;

                bool targetSheetOk = false;
                bool otherSheetsOk = true;

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
                    }
                    else
                    {
                        // 他のシートは「横向き」になっていたらNG（＝縦向きであるべき）
                        if (orientation == XlPageOrientation.xlLandscape)
                        {
                            otherSheetsOk = false;
                        }
                    }
                }

                // 両方の条件を満たしている場合のみ正解
                return targetSheetOk && otherSheetsOk;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (workbook != null)
                {
                    try { Marshal.ReleaseComObject(workbook); } catch { }
                }
            }
        }

        private bool CheckTask_2_1_02_Impl(string filePath)
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
                    excelApp = new Application();
                    excelApp.Visible = false;
                }
                
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
                
                if (workbook == null) return false;

                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;
                
                // 現在の印刷範囲を取得
                string currentPrintArea = worksheet.PageSetup.PrintArea;
                
                // 期待される印刷範囲をチェック（A6:G176）- 模試②用：解答手順あり模擬試験①問題文.csvを参照
                if (string.IsNullOrEmpty(currentPrintArea))
                {
                    return false;
                }
                
                // 印刷範囲を正規化（スペースを除去し、大文字に変換）
                string normalizedPrintArea = currentPrintArea.Replace(" ", "").ToUpper();
                
                // 期待される範囲A6:G176をチェック（模試②用）
                return normalizedPrintArea == "A6:G176" || normalizedPrintArea == "$A$6:$G$176";
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_2_1_03_Impl(string filePath)
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
                    excelApp = new Application();
                    excelApp.Visible = false;
                }
                
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
                
                if (workbook == null) return false;
                
                // 売上一覧シートを検索
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null)
                {
                    return false;
                }
                
                // 印刷タイトル（タイトル行）の設定をチェック - 6行目（模試②用：解答手順あり模擬試験①問題文.csvを参照）
                string titleRows = worksheet.PageSetup.PrintTitleRows;
                
                if (string.IsNullOrEmpty(titleRows))
                {
                    return false;
                }
                
                string normalizedTitleRows = titleRows.Replace("$", "").Replace(" ", "").ToUpper();
                return normalizedTitleRows == "6:6";
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_2_1_04_Impl(string filePath)
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
                    excelApp = new Application();
                    excelApp.Visible = false;
                }
                
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
                
                if (workbook == null) return false;
                
                // 販売実績シートを検索
                worksheet = FindWorksheet(workbook, "販売実績");
                if (worksheet == null)
                {
                    return false;
                }
                
                // 余白設定をチェック（「広い」の設定値）
                // 広い余白: 上下1インチ(72ポイント)、左右1インチ(72ポイント)
                double topMargin = worksheet.PageSetup.TopMargin;
                double bottomMargin = worksheet.PageSetup.BottomMargin;
                double leftMargin = worksheet.PageSetup.LeftMargin;
                double rightMargin = worksheet.PageSetup.RightMargin;
                
                // 72ポイント（1インチ）の許容範囲をチェック
                const double expectedMargin = 72.0;
                const double tolerance = 1.0;
                
                return Math.Abs(topMargin - expectedMargin) <= tolerance &&
                    Math.Abs(bottomMargin - expectedMargin) <= tolerance &&
                    Math.Abs(leftMargin - expectedMargin) <= tolerance &&
                    Math.Abs(rightMargin - expectedMargin) <= tolerance;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_2_1_05_Impl(string filePath)
        {
            // CSVの解答手順: 「販売実績」シート内の表を1ページ、グラフを2ページに印刷されるように改ページ設定
            // ExcelChecker1_1のCheckTask_1_1_05を参考に、改ページ位置をチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = false;
                }
                
                workbook = null;
                string fileName = Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "販売実績");
                if (worksheet == null) return false;
                
                // 改ページ位置をチェック（表が1ページ目、グラフが2ページ目になるように）
                // 模試②用：表示タブ＞改ページプレビューでページ線をドラッグ（通常、表の下に改ページが設定される）
                var hPageBreaks = worksheet.HPageBreaks;
                bool hasCorrectPageBreak = false;
                
                foreach (HPageBreak pageBreak in hPageBreaks)
                {
                    // 手動で設定された改ページのみをチェック
                    if (pageBreak.Type == XlPageBreak.xlPageBreakManual)
                    {
                        // 表とグラフの間に改ページがあることを確認
                        // 通常、表の下（例：行20付近）に改ページが設定される
                        if (pageBreak.Location.Row >= 15 && pageBreak.Location.Row <= 30)
                        {
                            hasCorrectPageBreak = true;
                            break;
                        }
                    }
                }
                
                return hasCorrectPageBreak;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_2_1_06_Impl(string filePath)
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
                    excelApp = new Application();
                    excelApp.Visible = false;
                }
                
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
                
                if (workbook == null) return false;
                
                // 模試結果シートを検索（模試②用：解答手順あり模擬試験①問題文.csvを参照）
                worksheet = FindWorksheet(workbook, "模試結果");
                if (worksheet == null)
                {
                    return false;
                }
                
                // テーブルの1行目（タイトル行）の折り返し設定をチェック
                // 模試②用：表の1行目（通常B7:L7などがタイトル行）
                Range targetRange = null;
                try
                {
                    targetRange = worksheet.Range["B7:L7"];
                    foreach (Range cell in targetRange.Cells)
                    {
                        if (!Convert.ToBoolean(cell.WrapText))
                        {
                            return false;
                        }
                    }
                    return true;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    if (targetRange != null) Marshal.ReleaseComObject(targetRange);
                }
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_2_1_07_Impl(string filePath)
        {
            // ExcelChecker1_1のCheckTask_1_1_07を参考に、G6以外にメモが存在しないかもチェック
            // 模試②用：解答手順あり模擬試験①問題文.csvを参照（G6に「タグを表示してください。」）
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
                    excelApp = new Application();
                    excelApp.Visible = false;
                }
                
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
                
                if (workbook == null) return false;
                
                // 売上一覧シートを検索
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null)
                {
                    return false;
                }
                
                // G6セルのメモをチェック（模試②用：解答手順あり模擬試験①問題文.csvを参照）
                targetCell = worksheet.Range["G6"];
                targetComment = targetCell.Comment;
                
                // セルにメモが存在するかチェック
                if (targetComment == null)
                {
                    return false;
                }
                
                // メモの内容が「タグを表示してください。」かチェック（模試②用）
                string commentText = targetComment.Text() ?? string.Empty;
                string normalizedText = commentText.Replace("\r", "").Replace("\n", "");
                int colonIndex = normalizedText.IndexOf(':');
                if (colonIndex >= 0 && colonIndex < normalizedText.Length - 1)
                {
                    normalizedText = normalizedText.Substring(colonIndex + 1);
                }
                if (!string.Equals(normalizedText.Trim(), "タグを表示してください。", StringComparison.Ordinal))
                {
                    return false;
                }

                // G6以外にメモが存在しないかチェック
                comments = worksheet.Comments;
                if (comments == null || comments.Count != 1)
                {
                    return false;
                }

                int commentCount = comments.Count;
                for (int i = 1; i <= commentCount; i++)
                {
                    Comment comment = null;
                    Range parentRange = null;
                    try
                    {
                        comment = comments.Item(i);
                        parentRange = comment.Parent as Range;
                        string address = parentRange?.Address[false, false] ?? string.Empty;

                        if (!string.Equals(address, "G6", StringComparison.OrdinalIgnoreCase))
                        {
                            return false;
                        }
                    }
                    finally
                    {
                        if (parentRange != null)
                        {
                            Marshal.ReleaseComObject(parentRange);
                        }
                        if (comment != null)
                        {
                            Marshal.ReleaseComObject(comment);
                        }
                    }
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (targetComment != null) Marshal.ReleaseComObject(targetComment);
                if (targetCell != null) Marshal.ReleaseComObject(targetCell);
                if (comments != null) Marshal.ReleaseComObject(comments);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private string GetCurrentExcelFilePath()
        {
            try
            {
                Application excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
            }
            catch (Exception)
            {
                // Excel アプリケーションが見つからない場合
            }
            return string.Empty;
        }

        private Worksheet FindWorksheet(Workbook workbook, string sheetName)
        {
            try
            {
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return worksheet;
                    }
                }
            }
            catch (Exception)
            {
                // シート検索でエラーが発生した場合
            }
            return null;
        }

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク01）
        /// </summary>
        private bool CheckTask_2_1_01_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                // config.jsonからパスを取得
                string initialFilePath = GetInitialDataFilePath(2, 1);
                string completedFilePath = GetCompletedDataFilePath(2, 1);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    // パスが取得できない場合は比較チェックをスキップ（既存チェックのみ）
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "売上一覧");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "売上一覧");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                // 印刷方向を比較
                return currentWorksheet.PageSetup.Orientation == completedWorksheet.PageSetup.Orientation;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComparisonComObject(completedWorksheet);
                ReleaseComparisonComObject(currentWorksheet);
                CloseComparisonWorkbook(completedWorkbook);
            }
        }

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク02）
        /// </summary>
        private bool CheckTask_2_1_02_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 1);
                string completedFilePath = GetCompletedDataFilePath(2, 1);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "売上一覧");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "売上一覧");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                // 印刷範囲を比較
                string currentPrintArea = currentWorksheet.PageSetup.PrintArea ?? string.Empty;
                string completedPrintArea = completedWorksheet.PageSetup.PrintArea ?? string.Empty;

                string normalizedCurrent = NormalizeRange(currentPrintArea);
                string normalizedCompleted = NormalizeRange(completedPrintArea);

                return normalizedCurrent == normalizedCompleted;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComparisonComObject(completedWorksheet);
                ReleaseComparisonComObject(currentWorksheet);
                CloseComparisonWorkbook(completedWorkbook);
            }
        }

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク03）
        /// </summary>
        private bool CheckTask_2_1_03_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 1);
                string completedFilePath = GetCompletedDataFilePath(2, 1);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "売上一覧");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "売上一覧");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                // 印刷タイトル行を比較
                string currentTitleRows = currentWorksheet.PageSetup.PrintTitleRows ?? string.Empty;
                string completedTitleRows = completedWorksheet.PageSetup.PrintTitleRows ?? string.Empty;

                string normalizedCurrent = NormalizeRange(currentTitleRows);
                string normalizedCompleted = NormalizeRange(completedTitleRows);

                return normalizedCurrent == normalizedCompleted;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComparisonComObject(completedWorksheet);
                ReleaseComparisonComObject(currentWorksheet);
                CloseComparisonWorkbook(completedWorkbook);
            }
        }

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク04）
        /// </summary>
        private bool CheckTask_2_1_04_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 1);
                string completedFilePath = GetCompletedDataFilePath(2, 1);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "販売実績");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "販売実績");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                // ページ設定を比較（余白を含む）
                return ComparePageSetup(currentWorksheet.PageSetup, completedWorksheet.PageSetup);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComparisonComObject(completedWorksheet);
                ReleaseComparisonComObject(currentWorksheet);
                CloseComparisonWorkbook(completedWorkbook);
            }
        }

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク05）
        /// </summary>
        private bool CheckTask_2_1_05_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 1);
                string completedFilePath = GetCompletedDataFilePath(2, 1);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "販売実績");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "販売実績");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                // ページ設定を比較（余白を含む）
                return ComparePageSetup(currentWorksheet.PageSetup, completedWorksheet.PageSetup);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComparisonComObject(completedWorksheet);
                ReleaseComparisonComObject(currentWorksheet);
                CloseComparisonWorkbook(completedWorkbook);
            }
        }

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク06）
        /// </summary>
        private bool CheckTask_2_1_06_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;
            Range currentRange = null;
            Range completedRange = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 1);
                string completedFilePath = GetCompletedDataFilePath(2, 1);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "模試結果");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "模試結果");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                // B7:L7の範囲の折り返し設定を比較（模試②用）
                currentRange = currentWorksheet.Range["B7:L7"];
                completedRange = completedWorksheet.Range["B7:L7"];

                if (currentRange == null || completedRange == null)
                    return false;

                // 各セルの折り返し設定を比較
                foreach (Range currentCell in currentRange.Cells)
                {
                    int row = currentCell.Row;
                    int col = currentCell.Column;
                    Range completedCell = completedWorksheet.Cells[row, col];

                    if (currentCell.WrapText != completedCell.WrapText)
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComparisonComObject(completedRange);
                ReleaseComparisonComObject(currentRange);
                ReleaseComparisonComObject(completedWorksheet);
                ReleaseComparisonComObject(currentWorksheet);
                CloseComparisonWorkbook(completedWorkbook);
            }
        }

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク07）
        /// </summary>
        private bool CheckTask_2_1_07_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;
            Range currentCell = null;
            Range completedCell = null;
            Comment currentComment = null;
            Comment completedComment = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 1);
                string completedFilePath = GetCompletedDataFilePath(2, 1);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "売上一覧");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "売上一覧");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                // G6セルのメモを比較（模試②用：解答手順あり模擬試験①問題文.csvを参照）
                currentCell = currentWorksheet.Range["G6"];
                completedCell = completedWorksheet.Range["G6"];

                if (currentCell == null || completedCell == null)
                    return false;

                currentComment = currentCell.Comment;
                completedComment = completedCell.Comment;

                // 両方にメモがあるか、両方にメモがないか
                if ((currentComment == null) != (completedComment == null))
                    return false;

                if (currentComment == null && completedComment == null)
                    return true;

                // メモの内容を比較
                string currentText = currentComment.Text() ?? string.Empty;
                string completedText = completedComment.Text() ?? string.Empty;

                // 正規化（改行やコロン以降の部分を除去）
                string normalizedCurrent = NormalizeCommentText(currentText);
                string normalizedCompleted = NormalizeCommentText(completedText);

                return normalizedCurrent == normalizedCompleted;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComparisonComObject(completedComment);
                ReleaseComparisonComObject(currentComment);
                ReleaseComparisonComObject(completedCell);
                ReleaseComparisonComObject(currentCell);
                ReleaseComparisonComObject(completedWorksheet);
                ReleaseComparisonComObject(currentWorksheet);
                CloseComparisonWorkbook(completedWorkbook);
            }
        }

        private string NormalizeCommentText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            
            string normalized = text.Replace("\r", "").Replace("\n", "");
            int colonIndex = normalized.IndexOf(':');
            if (colonIndex >= 0 && colonIndex < normalized.Length - 1)
            {
                normalized = normalized.Substring(colonIndex + 1);
            }
            return normalized.Trim();
        }

        // ExcelCheckerComparisonHelperの機能を直接実装
        private static JObject _config;

        static ExcelChecker2_1()
        {
            LoadConfig();
        }

        private static void LoadConfig()
        {
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    _config = JObject.Parse(json);
                }
            }
            catch
            {
                // config.jsonの読み込みに失敗した場合
            }
        }

        private static string GetInitialDataFilePath(int tabId, int projectId)
        {
            if (_config == null) return null;
            try
            {
                var projectData = _config["tabs"]?[tabId.ToString()]?["projects"]?[projectId.ToString()];
                return projectData?["initialDataFile"]?.ToString();
            }
            catch { return null; }
        }

        private static string GetCompletedDataFilePath(int tabId, int projectId)
        {
            if (_config == null) return null;
            try
            {
                var projectData = _config["tabs"]?[tabId.ToString()]?["projects"]?[projectId.ToString()];
                return projectData?["completedDataFile"]?.ToString();
            }
            catch { return null; }
        }

        private static Workbook GetComparisonWorkbook(Application excelApp, string filePath)
        {
            if (excelApp == null || string.IsNullOrEmpty(filePath)) return null;
            try
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
                if (File.Exists(filePath))
                {
                    return excelApp.Workbooks.Open(filePath, ReadOnly: true);
                }
            }
            catch { }
            return null;
        }

        private static Worksheet FindComparisonWorksheet(Workbook workbook, string sheetName)
        {
            if (workbook == null) return null;
            try
            {
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return worksheet;
                    }
                }
            }
            catch { }
            return null;
        }

        private static bool ComparePageSetup(PageSetup setup1, PageSetup setup2)
        {
            if (setup1 == null || setup2 == null) return false;
            try
            {
                if (setup1.Orientation != setup2.Orientation) return false;
                string printArea1 = setup1.PrintArea ?? string.Empty;
                string printArea2 = setup2.PrintArea ?? string.Empty;
                if (NormalizeRange(printArea1) != NormalizeRange(printArea2)) return false;
                string titleRows1 = setup1.PrintTitleRows ?? string.Empty;
                string titleRows2 = setup2.PrintTitleRows ?? string.Empty;
                if (NormalizeRange(titleRows1) != NormalizeRange(titleRows2)) return false;
                const double tolerance = 1.0;
                if (Math.Abs(setup1.TopMargin - setup2.TopMargin) > tolerance) return false;
                if (Math.Abs(setup1.BottomMargin - setup2.BottomMargin) > tolerance) return false;
                if (Math.Abs(setup1.LeftMargin - setup2.LeftMargin) > tolerance) return false;
                if (Math.Abs(setup1.RightMargin - setup2.RightMargin) > tolerance) return false;
                return true;
            }
            catch { return false; }
        }

        private static string NormalizeRange(string range)
        {
            if (string.IsNullOrEmpty(range)) return string.Empty;
            return range.Replace("$", "").Replace(" ", "").ToUpper();
        }

        private static void CloseComparisonWorkbook(Workbook workbook, bool saveChanges = false)
        {
            if (workbook != null)
            {
                try
                {
                    workbook.Close(saveChanges);
                    Marshal.ReleaseComObject(workbook);
                }
                catch { }
            }
        }

        private static void ReleaseComparisonComObject(object obj)
        {
            if (obj != null)
            {
                try { Marshal.ReleaseComObject(obj); }
                catch { }
            }
        }
    }
}
