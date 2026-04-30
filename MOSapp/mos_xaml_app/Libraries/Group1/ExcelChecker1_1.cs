using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;

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
                {
                    return false;
                }
                return CheckTask_1_1_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_1_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_1_02(filePath);
            }
            catch (Exception)
            {
                return false;
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
                {
                    return false;
                }
                return CheckTask_1_1_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_1_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_1_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_1_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_1_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_1_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_1_06(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_1_07()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_1_07(filePath);
            }
            catch (Exception)
            {
                return false;
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
                            otherSheetsOk = false;
                            System.Diagnostics.Debug.WriteLine($"[Task 1-1] Other sheet '{ws.Name}' is Landscape (Should be Portrait).");
                        }
                    }
                }

                // 両方の条件を満たしている場合のみ正解
                return targetSheetOk && otherSheetsOk;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Task 1-1] Error: {ex.Message}");
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
                    return false;
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
                    return false;
                }

                // 売上一覧シートを検索
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null)
                {
                    return false;
                }
                
                // 現在の印刷範囲を取得
                string currentPrintArea = worksheet.PageSetup.PrintArea;
                
                // 期待される印刷範囲をチェック（A4:F131）
                if (string.IsNullOrEmpty(currentPrintArea))
                {
                    return false;
                }
                
                // 印刷範囲を正規化（スペースを除去し、大文字に変換）
                string normalizedPrintArea = currentPrintArea.Replace(" ", "").ToUpper();
                
                // 期待される範囲A4:F131をチェック
                return normalizedPrintArea == "A4:F131" || normalizedPrintArea == "$A$4:$F$131";
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
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
                    return false;
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
                    return false;
                }

                // 売上一覧シートを検索
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null)
                {
                    return false;
                }
                
                // 印刷タイトル（タイトル行）の設定をチェック
                string titleRows = worksheet.PageSetup.PrintTitleRows;
                
                // 期待される設定: $2:$4 または 2:4
                if (string.IsNullOrEmpty(titleRows))
                {
                    return false;
                }
                
                string normalizedTitleRows = titleRows.Replace("$", "").Replace(" ", "").ToUpper();
                return normalizedTitleRows == "2:4";

            }
            catch (Exception)
            {
                return false;
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
                    return false;
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
                    return false;
                }

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
                    return false;
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
                    return false;
                }

                worksheet = FindWorksheet(workbook, "販売実績");
                if (worksheet == null)
                {
                    return false;
                }
                
                // 改ページ位置をチェック
                var vPageBreaks = worksheet.VPageBreaks;
                bool hasCorrectPageBreak = false;
                foreach (VPageBreak pageBreak in vPageBreaks)
                {
                    // 手動で設定された改ページのみをチェック
                    if (pageBreak.Location.Column == 8 && pageBreak.Type == XlPageBreak.xlPageBreakManual)
                    {
                        hasCorrectPageBreak = true;
                        break;
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
                    return false;
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
                    return false;
                }
                // スキルアップ検定結果シートを検索
                worksheet = FindWorksheet(workbook, "スキルアップ検定結果");
                if (worksheet == null)
                {
                    return false;
                }
                
                // A4:K4範囲の「折り返して全体を表示する」設定をチェック
                Range targetRange = worksheet.Range["A4:K4"];
                
                // 範囲内のすべてのセルで「折り返して全体を表示する」が設定されているかチェック
                foreach (Range cell in targetRange.Cells)
                {
                    if (!Convert.ToBoolean(cell.WrapText))
                    {
                        return false; // 一つでも設定されていないセルがあればfalse
                    }
                }
                
                return true; // すべてのセルで設定されていればtrue
            }
            catch (Exception)
            {
                return false;
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
                    return false;
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
                    return false;
                }

                // 売上一覧シートを検索
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null)
                {
                    return false;
                }
                
                // G4セルのメモをチェック
                targetCell = worksheet.Range["G4"];
                targetComment = targetCell.Comment;
                
                // セルにメモが存在するかチェック
                if (targetComment == null)
                {
                    return false;
                }
                
                // メモの内容が「最新の商品情報」かチェック
                string commentText = targetComment.Text() ?? string.Empty;
                string normalizedText = commentText.Replace("\r", "").Replace("\n", "");
                int colonIndex = normalizedText.IndexOf(':');
                if (colonIndex >= 0 && colonIndex < normalizedText.Length - 1)
                {
                    normalizedText = normalizedText.Substring(colonIndex + 1);
                }
                if (!string.Equals(normalizedText, "最新の商品情報", StringComparison.Ordinal))
                {
                    return false;
                }

                // G4以外にメモが存在しないかチェック
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

                        if (!string.Equals(address, "G4", StringComparison.OrdinalIgnoreCase))
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