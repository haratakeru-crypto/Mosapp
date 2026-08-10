using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group1
{
    public class ExcelChecker1_7
    {
        private static readonly string TARGET_FILE_PATH =
            MOSExcelMogiApp.Infrastructure.DataPathHelper.GetWorkingFilePath(1, 7);

        // Public wrappers
        public bool CheckTask_1_7_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_7_01_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_7_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_7_02_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_7_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_7_03_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_7_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_7_04_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_7_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_7_05_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_7_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_7_06_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_1_7_07()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_7_07_Impl(filePath);
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
                
                bool task1 = CheckTask_1_7_01_Impl(TARGET_FILE_PATH);
                results.Add($"Task 7-1 (数式コピー): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckTask_1_7_02_Impl(TARGET_FILE_PATH);
                results.Add($"Task 7-2 (MAX関数): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckTask_1_7_03_Impl(TARGET_FILE_PATH);
                results.Add($"Task 7-3 (COUNT関数): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckTask_1_7_04_Impl(TARGET_FILE_PATH);
                results.Add($"Task 7-4 (COUNTBLANK関数): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckTask_1_7_05_Impl(TARGET_FILE_PATH);
                results.Add($"Task 7-5 (RANDBETWEEN関数): {(task5 ? "OK" : "NG")}");
                
                bool task6 = CheckTask_1_7_06_Impl(TARGET_FILE_PATH);
                results.Add($"Task 7-6 (LEFT関数): {(task6 ? "OK" : "NG")}");
                
                bool task7 = CheckTask_1_7_07_Impl(TARGET_FILE_PATH);
                results.Add($"Task 7-7 (UNIQUE関数): {(task7 ? "OK" : "NG")}");
                
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
        private bool CheckTask_1_7_01_Impl(string filePath)
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
                    excelApp.Visible = false;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "イベント売上");
                if (worksheet == null) return false;
                
                // H5からH16までの範囲で数式がコピーされているかチェック
                Range h5Cell = worksheet.Range["H5"];
                Range targetRange = worksheet.Range["H6:H16"];
                
                bool h5HasFormula = h5Cell.HasFormula is bool && (bool)h5Cell.HasFormula;
                if (h5HasFormula)
                {
                    foreach (Range cell in targetRange.Cells)
                    {
                        bool cellHasFormula = cell.HasFormula is bool && (bool)cell.HasFormula;
                        if (!cellHasFormula)
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

        private bool CheckTask_1_7_02_Impl(string filePath)
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
                    excelApp.Visible = false;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "イベント売上");
                if (worksheet == null) return false;
                
                // J3セルのMAX関数をチェック
                Range targetCell = worksheet.Range["J3"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    
                    // デバッグ情報を出力
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 2 - Original formula: '{formula}'");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 2 - Normalized formula: '{normalizedFormula}'");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 2 - Contains MAX(: {normalizedFormula.Contains("MAX(")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 2 - Contains D5:G16: {normalizedFormula.Contains("D5:G16")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 2 - Contains $: {normalizedFormula.Contains("$")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 2 - Not contains $: {!normalizedFormula.Contains("$")}");
                    
                    // MAX関数と適切な範囲の組み合わせをチェック（テーブル参照も許容）
                    bool result = normalizedFormula.Contains("MAX(") && 
                                  (normalizedFormula.Contains("D5:G16") ||           // 通常の範囲参照
                                   normalizedFormula.Contains("売上一覧") ||            // テーブル名
                                   normalizedFormula.Contains("[[1日目]:[4日目]]")) &&  // テーブル列範囲
                                  !normalizedFormula.Contains("$");
                    
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 2 - Final result: {result}");
                    return result;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 2 - Formula is null");
                    return false;
                }
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

        private bool CheckTask_1_7_03_Impl(string filePath)
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
                    excelApp.Visible = false;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "試験結果");
                if (worksheet == null) return false;
                
                // I4セルのCOUNT関数をチェック
                Range targetCell = worksheet.Range["I4"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // COUNT関数と適切な範囲の組み合わせをチェック（テーブル参照も許容）
                    return normalizedFormula.Contains("COUNT(") && 
                           (normalizedFormula.Contains("L7:L56") ||           // 通常の範囲参照
                            normalizedFormula.Contains("試験結果") ||               // テーブル名
                            normalizedFormula.Contains("[[合計点]:[合計点]]")) &&   // テーブル列範囲
                           !normalizedFormula.Contains("$");
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

        private bool CheckTask_1_7_04_Impl(string filePath)
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
                    excelApp.Visible = false;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "試験結果");
                if (worksheet == null) return false;
                
                // J4セルのCOUNTBLANK関数をチェック
                Range targetCell = worksheet.Range["J4"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // COUNTBLANK関数と適切な範囲の組み合わせをチェック（テーブル参照も許容）
                    return normalizedFormula.Contains("COUNTBLANK(") && 
                           (normalizedFormula.Contains("L7:L56") ||               // 通常の範囲参照
                            normalizedFormula.Contains("試験結果") ||               // テーブル名
                            normalizedFormula.Contains("[[合計点]:[合計点]]")) &&   // テーブル列範囲
                           !normalizedFormula.Contains("$");
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

        private bool CheckTask_1_7_05_Impl(string filePath)
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
                    excelApp.Visible = false;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "試験結果");
                if (worksheet == null) return false;
                
                // B7からB56の範囲でRANDBETWEEN関数をチェック
                Range targetCell = worksheet.Range["B7"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // RANDBETWEEN関数とパラメータ1,8の組み合わせをチェック
                    return normalizedFormula.Contains("RANDBETWEEN(") && 
                           normalizedFormula.Contains("1,8");
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

        private bool CheckTask_1_7_06_Impl(string filePath)
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
                    excelApp.Visible = false;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "試験結果");
                if (worksheet == null) return false;
                
                // G7セルのLEFT関数をチェック
                Range targetCell = worksheet.Range["G7"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // LEFT関数と相対参照パラメータC7,2の組み合わせをチェック
                    return normalizedFormula.Contains("LEFT(") && 
                           normalizedFormula.Contains("C7,2") &&
                           !normalizedFormula.Contains("$");
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

        private bool CheckTask_1_7_07_Impl(string filePath)
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
                    excelApp.Visible = false;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "申込一覧");
                if (worksheet == null) return false;
                
                // I5セルのUNIQUE関数をチェック
                Range targetCell = worksheet.Range["I5"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    
                    // デバッグ情報を出力
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 7 - Original formula: '{formula}'");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 7 - Normalized formula: '{normalizedFormula}'");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 7 - Contains UNIQUE(: {normalizedFormula.Contains("UNIQUE(")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 7 - Contains E5:E174: {normalizedFormula.Contains("E5:E174")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 7 - Contains $: {normalizedFormula.Contains("$")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 7 - Not contains $: {!normalizedFormula.Contains("$")}");
                    
                    // UNIQUE関数と適切な範囲の組み合わせをチェック（テーブル参照も許容）
                    bool result = normalizedFormula.Contains("UNIQUE(") && 
                                  (normalizedFormula.Contains("E5:E174") ||           // 通常の範囲参照
                                   normalizedFormula.Contains("申込一覧") ||            // テーブル名
                                   normalizedFormula.Contains("[[学部]:[学部]]")) &&   // テーブル列範囲
                                  !normalizedFormula.Contains("$");
                    
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 7 - Final result: {result}");
                    return result;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 7 - Formula is null");
                    return false;
                }
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