using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group3
{
    public class ExcelChecker3_11
    {
        // Public CheckTask methods
        public bool CheckTask_3_11_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_11_01_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_11_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_11_02_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_11_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_11_03_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_11_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_11_04_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_11_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_11_05_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        // Private implementation methods
        private bool CheckTask_3_11_01_Impl(string filePath)
        {
            // 要件書: テストシートのセルA1に、「テスト」というコメントを挿入します。
            // ExcelChecker1_1のCheckTask_1_1_07を参考に、コメントをチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range targetCell = null;
            Comment targetComment = null;
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
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "テスト");
                if (worksheet == null) return false;
                
                // A1セルのコメントをチェック
                targetCell = worksheet.Range["A1"];
                targetComment = targetCell.Comment;
                
                if (targetComment == null) return false;
                
                string commentText = targetComment.Text() ?? string.Empty;
                string normalizedText = commentText.Replace("\r", "").Replace("\n", "");
                int colonIndex = normalizedText.IndexOf(':');
                if (colonIndex >= 0 && colonIndex < normalizedText.Length - 1)
                {
                    normalizedText = normalizedText.Substring(colonIndex + 1);
                }
                
                return string.Equals(normalizedText.Trim(), "テスト", StringComparison.Ordinal);
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (targetComment != null) Marshal.ReleaseComObject(targetComment);
                if (targetCell != null) Marshal.ReleaseComObject(targetCell);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_11_02_Impl(string filePath)
        {
            // 要件書: テストシートのセルA1のコメントを表示したままにします。
            // コメントが表示されているかチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range targetCell = null;
            Comment targetComment = null;
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
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "テスト");
                if (worksheet == null) return false;
                
                // A1セルのコメントが表示されているかチェック
                targetCell = worksheet.Range["A1"];
                targetComment = targetCell.Comment;
                
                if (targetComment == null) return false;
                
                // コメントが表示されているかチェック（Visibleプロパティ）
                bool isVisible = targetComment.Visible;
                return isVisible;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (targetComment != null) Marshal.ReleaseComObject(targetComment);
                if (targetCell != null) Marshal.ReleaseComObject(targetCell);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_11_03_Impl(string filePath)
        {
            // 要件書: テストシートのセルB1に、入力規則を設定します。整数で1から10までの値のみ入力できるようにします。
            // 入力規則をチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range targetCell = null;
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
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "テスト");
                if (worksheet == null) return false;
                
                // B1セルの入力規則をチェック
                targetCell = worksheet.Range["B1"];
                Validation validation = targetCell.Validation;
                
                if (validation == null) return false;
                
                // 入力規則の種類をチェック（整数）
                XlDVType type = validation.Type;
                if (type != XlDVType.xlValidateWholeNumber) return false;
                
                // 最小値と最大値をチェック（1から10まで）
                int minValue = (int)validation.Formula1;
                int maxValue = (int)validation.Formula2;
                
                return minValue == 1 && maxValue == 10;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (targetCell != null) Marshal.ReleaseComObject(targetCell);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_11_04_Impl(string filePath)
        {
            // 要件書: テストシートのセルB1に入力規則のエラーメッセージを設定します。タイトルは「エラー」、エラーメッセージは「1から10までの数値を入力してください」とします。
            // 入力規則のエラーメッセージをチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range targetCell = null;
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
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "テスト");
                if (worksheet == null) return false;
                
                // B1セルの入力規則のエラーメッセージをチェック
                targetCell = worksheet.Range["B1"];
                Validation validation = targetCell.Validation;
                
                if (validation == null) return false;
                
                // エラーメッセージのタイトルとメッセージをチェック
                string errorTitle = validation.ErrorTitle ?? "";
                string errorMessage = validation.ErrorMessage ?? "";
                
                bool titleCorrect = errorTitle.Contains("エラー");
                bool messageCorrect = errorMessage.Contains("1から10までの数値を入力してください");
                
                return titleCorrect && messageCorrect;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (targetCell != null) Marshal.ReleaseComObject(targetCell);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_11_05_Impl(string filePath)
        {
            // 要件書: テストシートのセルC1に、ウィンドウ枠の固定を設定します。
            // ExcelChecker2_6のCheckTask_2_6_01_Implを参考に、ウィンドウ枠の固定をチェック
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
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "テスト");
                if (worksheet == null) return false;
                
                // ウィンドウ枠の固定をチェック
                worksheet.Activate();
                Window window = excelApp.ActiveWindow;
                bool isFrozen = window.FreezePanes;
                
                // C1セルで固定されている場合、SplitRow >= 1 または SplitColumn >= 3
                if (isFrozen)
                {
                    // C1セルで固定されている場合、SplitColumn >= 3 または SplitRow >= 1
                    if (window.SplitColumn >= 3 || window.SplitRow >= 1)
                    {
                        return true;
                    }
                }
                
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        // Helper methods
        private string GetCurrentExcelFilePath()
        {
            Application excelApp = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
            }
            catch
            {
                // Excel is not running or no active workbook
            }
            finally
            {
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
            return string.Empty;
        }

        private Worksheet FindWorksheet(Workbook workbook, string worksheetName)
        {
            try
            {
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return worksheet;
                    }
                }
            }
            catch
            {
                // Error accessing worksheets
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

