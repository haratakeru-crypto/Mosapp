using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace mogiExcelChecker1
{
    public class mogiExcelChecker1
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\mogi_project1.xlsx";

        /// <summary>
        /// Project1のタスク1-1をチェックする
        /// 文化祭コピーシートの表にL8:Q16をコピーして貼り付け
        /// </summary>
        public bool CheckProject_01_Task_01_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_01_Task_01_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project1のタスク1-2をチェックする
        /// 文化祭切り取りシートの表にL8:Q16を切り取って貼り付け
        /// </summary>
        public bool CheckProject_01_Task_01_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_01_Task_01_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project1のタスク1-3をチェックする
        /// 文化祭切り取りシートのテーブルのH列の式オートフィル
        /// </summary>
        public bool CheckProject_01_Task_01_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_01_Task_01_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project1のタスク1-4をチェックする
        /// 担当者マスターシートの表にG11:J16をコピーして列幅を揃えて貼り付け
        /// </summary>
        public bool CheckProject_01_Task_01_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_01_Task_01_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project1のタスク1-5をチェックする
        /// 担当者マスターシートの売上にL11:L16をリンク貼り付け
        /// </summary>
        public bool CheckProject_01_Task_01_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_01_Task_01_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_1()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: mogi_project1.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: mogi_project1.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_01_Task_01_01(TARGET_FILE_PATH);
                results.Add($"Task 1-1 (コピー貼り付け): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_01_Task_01_02(TARGET_FILE_PATH);
                results.Add($"Task 1-2 (切り取り貼り付け): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_01_Task_01_03(TARGET_FILE_PATH);
                results.Add($"Task 1-3 (オートフィル): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_01_Task_01_04(TARGET_FILE_PATH);
                results.Add($"Task 1-4 (列幅保持貼り付け): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckProject_01_Task_01_05(TARGET_FILE_PATH);
                results.Add($"Task 1-5 (リンク貼り付け): {(task5 ? "OK" : "NG")}");
                
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

        private bool CheckProject_01_Task_01_01(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "文化祭コピー");
                if (worksheet == null) return false;
                
                // B8:G16の範囲にデータがコピーされているかチェック
                Range sourceRange = worksheet.Range["L8:Q16"];
                Range targetRange = worksheet.Range["B8:G16"];
                
                // 範囲内にデータが存在するかチェック
                for (int i = 1; i <= sourceRange.Rows.Count; i++)
                {
                    for (int j = 1; j <= sourceRange.Columns.Count; j++)
                    {
                        var sourceValue = sourceRange.Cells[i, j].Value2;
                        var targetValue = targetRange.Cells[i, j].Value2;
                        if (sourceValue != null && targetValue != null)
                        {
                            if (sourceValue.ToString() == targetValue.ToString())
                            {
                                return true; // 一部でもコピーされていればOK
                            }
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

        private bool CheckProject_01_Task_01_02(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "文化祭切り取り");
                if (worksheet == null) return false;
                
                // L8:Q16の範囲が空でB8:G16に移動しているかチェック
                Range sourceRange = worksheet.Range["L8:Q16"];
                Range targetRange = worksheet.Range["B8:G16"];
                
                // 元の場所が空かチェック
                bool sourceEmpty = true;
                foreach (Range cell in sourceRange.Cells)
                {
                    if (cell.Value2 != null)
                    {
                        sourceEmpty = false;
                        break;
                    }
                }
                
                // ターゲットにデータがあるかチェック
                bool targetHasData = false;
                foreach (Range cell in targetRange.Cells)
                {
                    if (cell.Value2 != null)
                    {
                        targetHasData = true;
                        break;
                    }
                }
                
                return sourceEmpty && targetHasData;
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

        private bool CheckProject_01_Task_01_03(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "文化祭切り取り");
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

        private bool CheckProject_01_Task_01_04(string filePath)
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
                
                // A11:D16の範囲にG11:J16がコピーされているかチェック
                Range sourceRange = worksheet.Range["G11:J16"];
                Range targetRange = worksheet.Range["A11:D16"];
                
                for (int i = 1; i <= sourceRange.Rows.Count; i++)
                {
                    for (int j = 1; j <= sourceRange.Columns.Count; j++)
                    {
                        var sourceValue = sourceRange.Cells[i, j].Value2;
                        var targetValue = targetRange.Cells[i, j].Value2;
                        if (sourceValue != null && targetValue != null)
                        {
                            if (sourceValue.ToString() == targetValue.ToString())
                            {
                                return true; // 一部でもコピーされていればOK
                            }
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

        private bool CheckProject_01_Task_01_05(string filePath)
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
                
                // E11:E16の範囲にL11:L16がリンク貼り付けされているかチェック
                Range targetRange = worksheet.Range["E11:E16"];
                
                foreach (Range cell in targetRange.Cells)
                {
                    if (cell.HasFormula)
                    {
                        string formula = cell.Formula;
                        // リンク数式が含まれているかチェック（L列への参照）
                        if (formula != null && formula.Contains("L"))
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
    }
} 