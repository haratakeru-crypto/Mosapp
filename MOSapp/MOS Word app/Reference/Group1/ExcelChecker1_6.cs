using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group1
{
    public class ExcelChecker1_6
    {
        private const string TARGET_FILE_PATH = @"C:\\MOSTest\\Excel365\\project6.xlsx";

        // Public wrappers
        public bool CheckTask_1_6_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_6_01_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_6_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_6_02_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_6_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_6_03_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_6_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_6_04_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_6()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: project6.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: project6.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                bool task1 = CheckTask_1_6_01_Impl(TARGET_FILE_PATH);
                results.Add($"Task 6-1 (ウィンドウ枠の固定): {(task1 ? "OK" : "NG")}");

                bool task2 = CheckTask_1_6_02_Impl(TARGET_FILE_PATH);
                results.Add($"Task 6-2 (ハイパーリンク): {(task2 ? "OK" : "NG")}");

                bool task3 = CheckTask_1_6_03_Impl(TARGET_FILE_PATH);
                results.Add($"Task 6-3 (通貨書式): {(task3 ? "OK" : "NG")}");

                bool task4 = CheckTask_1_6_04_Impl(TARGET_FILE_PATH);
                results.Add($"Task 6-4 (プロパティタグ): {(task4 ? "OK" : "NG")}");

                return string.Join("\n", results);
            }
            catch (Exception ex)
            {
                return $"エラー: {ex.Message}";
            }
        }

        // Helpers
        private bool IsExcelFileOpen(string filePath)
        {
            Application excelApp = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                foreach (Workbook workbook in excelApp.Workbooks)
                {
                    if (string.Equals(workbook.FullName, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);
            }
        }

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
                return null;
            }
            catch (COMException)
            {
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

        // Private implementations
        private bool CheckTask_1_6_01_Impl(string filePath)
        {
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
                    excelApp.Visible = true;
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;

                worksheet.Activate();
                Window window = excelApp.ActiveWindow;
                if (window.FreezePanes)
                {
                    if (window.SplitRow >= 4)
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
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_6_02_Impl(string filePath)
        {
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
                    excelApp.Visible = true;
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;

                foreach (Hyperlink hyperlink in worksheet.Hyperlinks)
                {
                    if (hyperlink.Address != null &&
                        hyperlink.Address.Contains("https://rabbitway.jp/service_mos"))
                    {
                        if (hyperlink.TextToDisplay != null &&
                            hyperlink.TextToDisplay.Contains("パソコン資格講座のご相談"))
                        {
                            return true;
                        }
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
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_6_03_Impl(string filePath)
        {
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
                    excelApp.Visible = true;
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                worksheet = FindWorksheet(workbook, "販売実績");
                if (worksheet == null) return false;

                Range targetRange = worksheet.Range["B5:G11"];
                foreach (Range cell in targetRange.Cells)
                {
                    string numberFormat = cell.NumberFormat as string;
                    if (numberFormat == null) return false;
                    
                    // 通貨形式であることを確認（¥または$を含む）
                    bool isCurrency = numberFormat.Contains("¥") || numberFormat.Contains("$");
                    // 小数点が表示されていないことを確認
                    bool hasNoDecimals = !numberFormat.Contains(".0");
                    
                    // すべてのセルが通貨形式（小数点なし）である必要がある
                    if (!isCurrency || !hasNoDecimals)
                    {
                        return false;
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
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_6_04_Impl(string filePath)
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
                    excelApp = new Application();
                    excelApp.Visible = true;
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                try
                {
                    dynamic properties = workbook.BuiltinDocumentProperties;
                    dynamic keywordsProperty = properties["Keywords"];
                    string tags = keywordsProperty.Value as string;
                    if (tags != null && tags.Contains("売上"))
                    {
                        return true;
                    }
                }
                catch
                {
                    return false;
                }

                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                // workbookはCloseしない
            }
        }
    }
}