using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;

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
                return CheckTask_2_1_01_Impl(filePath);
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
                return CheckTask_2_1_02_Impl(filePath);
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
                return CheckTask_2_1_03_Impl(filePath);
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
                return CheckTask_2_1_04_Impl(filePath);
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
                return CheckTask_2_1_05_Impl(filePath);
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
                return CheckTask_2_1_06_Impl(filePath);
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
                return CheckTask_2_1_07_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_1_01_Impl(string filePath)
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
                    excelApp.Visible = true;
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
                
                // 印刷方向をチェック
                XlPageOrientation orientation = worksheet.PageSetup.Orientation;
                
                return orientation == XlPageOrientation.xlLandscape;
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
                    excelApp.Visible = true;
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
                
                // 期待される印刷範囲をチェック（A6:G176）
                if (string.IsNullOrEmpty(currentPrintArea))
                {
                    return false;
                }
                
                // 印刷範囲を正規化（スペースを除去し、大文字に変換）
                string normalizedPrintArea = currentPrintArea.Replace(" ", "").ToUpper();
                
                // 期待される範囲A6:G176をチェック
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
                    excelApp.Visible = true;
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
                
                // 印刷方向をチェック
                XlPageOrientation orientation = worksheet.PageSetup.Orientation;
                
                return orientation == XlPageOrientation.xlLandscape;
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
                    excelApp.Visible = true;
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
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_2_1_05_Impl(string filePath)
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
                    excelApp.Visible = true;
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
                    excelApp.Visible = true;
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
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_2_1_07_Impl(string filePath)
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
                    excelApp.Visible = true;
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
    }
}
