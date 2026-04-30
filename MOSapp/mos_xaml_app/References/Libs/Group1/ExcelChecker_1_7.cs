using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelChecker7
{
    public class ExcelChecker7
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\project7.xlsx";

        /// <summary>
        /// Project7のタスク7-1をチェックする
        /// シート[イベント売上]のセル【H5】の数式をコピーして、[4日間合計]の列を完成させる
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_07_Task_07_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_07_Task_07_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project7のタスク7-2をチェックする
        /// シート[イベント売上]のセル【J3】に、MAX関数を使って「最高売上」を算出
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_07_Task_07_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_07_Task_07_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project7のタスク7-3をチェックする
        /// シート[試験結果]のセル【I4】に、COUNT関数を使って「出席者」を算出
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_07_Task_07_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_07_Task_07_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project7のタスク7-4をチェックする
        /// シート[試験結果]のセル【J4】に、COUNTBLANK関数を使って「欠席者」を算出
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_07_Task_07_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_07_Task_07_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project7のタスク7-5をチェックする
        /// シート[試験結果]の「グループ」の列に、RANDBETWEEN関数を使ってランダムな番号を表示
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_07_Task_07_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_07_Task_07_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project7のタスク7-6をチェックする
        /// シート[試験結果]の「学部学科略称」の列に、LEFT関数を使って「学籍番号」の左端から2文字を表示
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_07_Task_07_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_07_Task_07_06(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project7のタスク7-7をチェックする
        /// シート[申込一覧]の「学部一覧」の列に、UNIQUE関数を使って「学部」を重複しないように表示
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_07_Task_07_07()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_07_Task_07_07(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_7()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: project7.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: project7.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_07_Task_07_01(TARGET_FILE_PATH);
                results.Add($"Task 7-1 (数式コピー): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_07_Task_07_02(TARGET_FILE_PATH);
                results.Add($"Task 7-2 (MAX関数): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_07_Task_07_03(TARGET_FILE_PATH);
                results.Add($"Task 7-3 (COUNT関数): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_07_Task_07_04(TARGET_FILE_PATH);
                results.Add($"Task 7-4 (COUNTBLANK関数): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckProject_07_Task_07_05(TARGET_FILE_PATH);
                results.Add($"Task 7-5 (RANDBETWEEN関数): {(task5 ? "OK" : "NG")}");
                
                bool task6 = CheckProject_07_Task_07_06(TARGET_FILE_PATH);
                results.Add($"Task 7-6 (LEFT関数): {(task6 ? "OK" : "NG")}");
                
                bool task7 = CheckProject_07_Task_07_07(TARGET_FILE_PATH);
                results.Add($"Task 7-7 (UNIQUE関数): {(task7 ? "OK" : "NG")}");
                
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

        private bool CheckProject_07_Task_07_01(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "イベント売上");
                if (worksheet == null) return false;
                
                // H5からH16までの範囲で数式がコピーされているかチェック
                Range h5Cell = worksheet.Range["H5"];
                Range targetRange = worksheet.Range["H6:H16"];
                
                if (h5Cell.HasFormula)
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

        private bool CheckProject_07_Task_07_02(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "イベント売上");
                if (worksheet == null) return false;
                
                // J3セルのMAX関数をチェック
                Range targetCell = worksheet.Range["J3"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("MAX(D5:G16)") || normalizedFormula.Contains("=MAX(D5:G16)");
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

        private bool CheckProject_07_Task_07_03(string filePath)
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
                
                // I4セルのCOUNT関数をチェック
                Range targetCell = worksheet.Range["I4"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("COUNT(L7:L56)") || normalizedFormula.Contains("=COUNT(L7:L56)");
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

        private bool CheckProject_07_Task_07_04(string filePath)
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
                
                // J4セルのCOUNTBLANK関数をチェック
                Range targetCell = worksheet.Range["J4"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("COUNTBLANK(L7:L56)") || normalizedFormula.Contains("=COUNTBLANK(L7:L56)");
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

        private bool CheckProject_07_Task_07_05(string filePath)
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
                
                // B7からB56の範囲でRANDBETWEEN関数をチェック
                Range targetCell = worksheet.Range["B7"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("RANDBETWEEN(1,8)") || normalizedFormula.Contains("=RANDBETWEEN(1,8)");
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

        private bool CheckProject_07_Task_07_06(string filePath)
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
                
                // G7セルのLEFT関数をチェック
                Range targetCell = worksheet.Range["G7"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("LEFT(C7,2)") || normalizedFormula.Contains("=LEFT(C7,2)");
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

        private bool CheckProject_07_Task_07_07(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "申込一覧");
                if (worksheet == null) return false;
                
                // I5セルのUNIQUE関数をチェック
                Range targetCell = worksheet.Range["I5"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("UNIQUE(E5:E174)") || normalizedFormula.Contains("=UNIQUE(E5:E174)");
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