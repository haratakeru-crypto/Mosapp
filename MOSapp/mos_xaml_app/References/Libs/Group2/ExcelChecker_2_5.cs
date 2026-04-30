using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace mogiExcelChecker5
{
    public class mogiExcelChecker5
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\mogi_project5.xlsx";

        /// <summary>
        /// Project5のタスク5-1をチェックする
        /// 試験結果シートのB5を基準にテキストファイル「結果内容」をインポートする
        /// </summary>
        public bool CheckProject_05_Task_05_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_05_Task_05_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project5のタスク5-2をチェックする
        /// テーブルのスタイルを「緑、テーブルスタイル（濃色）11」に変更
        /// </summary>
        public bool CheckProject_05_Task_05_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_05_Task_05_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project5のタスク5-3をチェックする
        /// テーブルの最後の列を強調
        /// </summary>
        public bool CheckProject_05_Task_05_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_05_Task_05_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project5のタスク5-4をチェックする
        /// テーブルの縞模様を解除
        /// </summary>
        public bool CheckProject_05_Task_05_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_05_Task_05_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project5のタスク5-5をチェックする
        /// 担当者マスターシートのテーブルを有楽町店の担当だけ表示
        /// </summary>
        public bool CheckProject_05_Task_05_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_05_Task_05_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project5のタスク5-6をチェックする
        /// 担当者マスターシートの数式を表示
        /// </summary>
        public bool CheckProject_05_Task_05_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_05_Task_05_06(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_5()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: mogi_project5.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: mogi_project5.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_05_Task_05_01(TARGET_FILE_PATH);
                results.Add($"Task 5-1 (テキストインポート): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_05_Task_05_02(TARGET_FILE_PATH);
                results.Add($"Task 5-2 (テーブルスタイル): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_05_Task_05_03(TARGET_FILE_PATH);
                results.Add($"Task 5-3 (最後の列強調): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_05_Task_05_04(TARGET_FILE_PATH);
                results.Add($"Task 5-4 (縞模様解除): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckProject_05_Task_05_05(TARGET_FILE_PATH);
                results.Add($"Task 5-5 (フィルタリング): {(task5 ? "OK" : "NG")}");
                
                bool task6 = CheckProject_05_Task_05_06(TARGET_FILE_PATH);
                results.Add($"Task 5-6 (数式表示): {(task6 ? "OK" : "NG")}");
                
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

        private bool CheckProject_05_Task_05_01(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "試験結果");
                if (worksheet == null) return false;
                
                // B5セルにテキストインポートされたデータが存在するかチェック
                Range targetCell = worksheet.Range["B5"];
                var value = targetCell.Value2;
                if (value != null)
                {
                    // インポートされたデータが存在することを確認
                    return !string.IsNullOrEmpty(value.ToString());
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

        private bool CheckProject_05_Task_05_02(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "試験結果");
                if (worksheet == null) return false;
                
                // テーブルスタイルが濃色11（緑）に設定されているかチェック
                ListObjects listObjects = worksheet.ListObjects;
                if (listObjects.Count > 0)
                {
                    ListObject table = listObjects[1];
                    string tablestyle = table.TableStyle;
                    return tablestyle != null && 
                           (tablestyle.Contains("TableStyleDark11") || 
                            tablestyle.Contains("Dark11"));
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

        private bool CheckProject_05_Task_05_03(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "試験結果");
                if (worksheet == null) return false;
                
                // テーブルの最後の列が強調されているかチェック
                ListObjects listObjects = worksheet.ListObjects;
                if (listObjects.Count > 0)
                {
                    ListObject table = listObjects[1];
                    return table.ShowTableStyleLastColumn;
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

        private bool CheckProject_05_Task_05_04(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "試験結果");
                if (worksheet == null) return false;
                
                // テーブルの縞模様が解除されているかチェック
                ListObjects listObjects = worksheet.ListObjects;
                if (listObjects.Count > 0)
                {
                    ListObject table = listObjects[1];
                    return !table.ShowTableStyleRowStripes && !table.ShowTableStyleColumnStripes;
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

        private bool CheckProject_05_Task_05_05(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "担当者マスター");
                if (worksheet == null) return false;
                
                // オートフィルタが適用されているかチェック
                return worksheet.AutoFilterMode;
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

        private bool CheckProject_05_Task_05_06(string filePath)
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
                
                // 数式が表示されているかチェック
                Window window = excelApp.ActiveWindow;
                if (window != null)
                {
                    return window.DisplayFormulas;
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