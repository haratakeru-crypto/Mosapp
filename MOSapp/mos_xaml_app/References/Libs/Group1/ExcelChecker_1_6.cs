using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelChecker6
{
    public class ExcelChecker6
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\project6.xlsx";

        /// <summary>
        /// Project6のタスク6-1をチェックする
        /// シート[売上一覧]の1～4行目が常に表示されるように設定
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_06_Task_06_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_06_Task_06_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project6のタスク6-2をチェックする
        /// シート[売上一覧]の文字列「パソコン資格講座のご相談」にハイパーリンクを挿入
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_06_Task_06_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_06_Task_06_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project6のタスク6-3をチェックする
        /// シート[販売実績]の「1月」から「6月」までの数値の書式を「通貨」に変更
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_06_Task_06_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_06_Task_06_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project6のタスク6-4をチェックする
        /// プロパティのタグに「売上」と追加
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_06_Task_06_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_06_Task_06_04(filePath);
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
                
                bool task1 = CheckProject_06_Task_06_01(TARGET_FILE_PATH);
                results.Add($"Task 6-1 (ウィンドウ枠の固定): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_06_Task_06_02(TARGET_FILE_PATH);
                results.Add($"Task 6-2 (ハイパーリンク): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_06_Task_06_03(TARGET_FILE_PATH);
                results.Add($"Task 6-3 (通貨書式): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_06_Task_06_04(TARGET_FILE_PATH);
                results.Add($"Task 6-4 (プロパティタグ): {(task4 ? "OK" : "NG")}");
                
                return string.Join("\n", results);
            }
            catch (Exception ex)
            {
                return $"エラー: {ex.Message}";
            }
        }

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
            string fileName = System.IO.Path.GetFileName(filePath);
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

        private bool CheckProject_06_Task_06_01(string filePath)
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
                
                // ウィンドウ枠の固定をチェック
                Window window = excelApp.ActiveWindow;
                if (window.FreezePanes)
                {
                    // 固定された行数をチェック
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

        private bool CheckProject_06_Task_06_02(string filePath)
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
                
                // ハイパーリンクを検索
                foreach (Hyperlink hyperlink in worksheet.Hyperlinks)
                {
                    if (hyperlink.Address != null && 
                        hyperlink.Address.Contains("https://rabbitway.jp/service_mos"))
                    {
                        // テキストもチェック
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

        private bool CheckProject_06_Task_06_03(string filePath)
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
                
                // セル範囲B5:G11の書式をチェック
                Range targetRange = worksheet.Range["B5:G11"];
                foreach (Range cell in targetRange.Cells)
                {
                    // 通貨書式かチェック
                    string numberFormat = cell.NumberFormat;
                    if (numberFormat.Contains("¥") || numberFormat.Contains("$") || 
                        numberFormat.Contains("Currency"))
                    {
                        // 小数点がないかチェック
                        if (!numberFormat.Contains(".0"))
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

        private bool CheckProject_06_Task_06_04(string filePath)
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
                
                // プロパティのタグをチェック
                try
                {
                    string tags = workbook.BuiltinDocumentProperties["Keywords"].Value;
                    if (tags != null && tags.Contains("売上"))
                    {
                        return true;
                    }
                }
                catch
                {
                    // BuiltinDocumentPropertiesでアクセスできない場合
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
                // workbookはCloseしない（既存のファイルを操作しているため）
            }
        }
    }
} 