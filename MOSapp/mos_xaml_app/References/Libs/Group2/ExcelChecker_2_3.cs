using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic; // Added for ListObjects and ListObject

namespace mogiExcelChecker3
{
    public class mogiExcelChecker3
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\mogi_project3.xlsx";

        /// <summary>
        /// Project3のタスク3-1をチェックする
        /// 担当者マスターシートのE列に勤続年数が5より大きければ「あり」、そうでなければ「なし」と表示
        /// </summary>
        public bool CheckProject_03_Task_03_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project3のタスク3-2をチェックする
        /// 表のアカウントとアドレス「@win.jp」を組み合わせてメールアドレスの列に表示
        /// </summary>
        public bool CheckProject_03_Task_03_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project3のタスク3-3をチェックする
        /// 「文化祭」シートのテーブルのH列の式を下まで完成させる
        /// </summary>
        public bool CheckProject_03_Task_03_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project3のタスク3-4をチェックする
        /// 「文化祭」シートのテーブルに4日間合計の列を追加
        /// </summary>
        public bool CheckProject_03_Task_03_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project3のタスク3-5をチェックする
        /// 「文化祭」シートの4日間の売上合計の最高金額をJ6に表示
        /// </summary>
        public bool CheckProject_03_Task_03_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_3()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: mogi_project3.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: mogi_project3.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_03_Task_03_01(TARGET_FILE_PATH);
                results.Add($"Task 3-1 (IF関数): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_03_Task_03_02(TARGET_FILE_PATH);
                results.Add($"Task 3-2 (CONCAT関数メール): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_03_Task_03_03(TARGET_FILE_PATH);
                results.Add($"Task 3-3 (H列オートフィル): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_03_Task_03_04(TARGET_FILE_PATH);
                results.Add($"Task 3-4 (テーブル列追加): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckProject_03_Task_03_05(TARGET_FILE_PATH);
                results.Add($"Task 3-5 (MAX関数): {(task5 ? "OK" : "NG")}");
                
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

        private bool CheckProject_03_Task_03_01(string filePath)
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
                
                // E11セルのIF関数をチェック
                Range targetCell = worksheet.Range["E11"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // IF関数でD11>5の条件をチェック
                    return normalizedFormula.Contains("IF(D11>5,\"あり\",\"なし\")") || 
                           normalizedFormula.Contains("=IF(D11>5,\"あり\",\"なし\")") ||
                           (normalizedFormula.Contains("IF") && normalizedFormula.Contains("D11>5"));
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

        private bool CheckProject_03_Task_03_02(string filePath)
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
                
                // F11セルのCONCAT関数をチェック
                Range targetCell = worksheet.Range["F11"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // CONCAT関数でC11と@win.jpを結合しているかチェック
                    return normalizedFormula.Contains("CONCAT(C11,\"@WIN.JP\")") || 
                           normalizedFormula.Contains("=CONCAT(C11,\"@WIN.JP\")") ||
                           (normalizedFormula.Contains("CONCAT") && normalizedFormula.Contains("C11") && normalizedFormula.Contains("@WIN.JP"));
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

        private bool CheckProject_03_Task_03_03(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "文化祭");
                if (worksheet == null) return false;
                
                // H8からH16までの範囲で数式がオートフィルされているかチェック
                Range h8Cell = worksheet.Range["H8"];
                Range targetRange = worksheet.Range["H9:H16"];
                
                if (h8Cell.HasFormula)
                {
                    foreach (Range cell in targetRange.Cells)
                    {
                        if (!cell.HasFormula)
                        {
                            return false; // 一つでも数式がないセルがあればfalse
                        }
                    }
                    return true;
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

        private bool CheckProject_03_Task_03_04(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "文化祭");
                if (worksheet == null) return false;
                
                // テーブルのサイズが変更されているかチェック
                ListObjects listObjects = worksheet.ListObjects;
                if (listObjects.Count > 0)
                {
                    ListObject table = listObjects[1];
                    // テーブルの範囲がB7:H16かチェック
                    string tableRange = table.Range.Address;
                    return tableRange.Contains("B7:H16") || tableRange.Contains("$B$7:$H$16");
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

        private bool CheckProject_03_Task_03_05(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "文化祭");
                if (worksheet == null) return false;
                
                // J6セルのMAX関数をチェック
                Range targetCell = worksheet.Range["J6"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // MAX関数またはオートSUMの最大値が設定されているかチェック
                    return normalizedFormula.Contains("MAX") ||
                           normalizedFormula.Contains("=MAX");
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
    }
} 