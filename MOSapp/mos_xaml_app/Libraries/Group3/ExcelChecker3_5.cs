using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group3
{
    public class ExcelChecker3_5
    {
        // Public CheckTask methods
        public bool CheckTask_3_5_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_5_01_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_5_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_5_02_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_5_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_5_03_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_5_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_5_04_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_5_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_5_05_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }


        // Private implementation methods
        private bool CheckTask_3_5_01_Impl(string filePath)
        {
            // 要件書: 受注一覧シートのJ列に、商品IDの左から2文字を表示します。
            // ExcelChecker2_7のCheckTask_2_7_05_Implを参考に、LEFT関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "受注一覧");
                if (worksheet == null) return false;
                
                // J4セルのLEFT関数をチェック
                Range targetCell = worksheet.Range["J4"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // LEFT関数とC4,2の組み合わせをチェック
                    return normalizedFormula.Contains("LEFT(") && 
                           normalizedFormula.Contains("C4,2") &&
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
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_5_02_Impl(string filePath)
        {
            // 要件書: 受注一覧シートのK列に、商品IDの3文字目から4文字を表示します。
            // MID関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "受注一覧");
                if (worksheet == null) return false;
                
                // K4セルのMID関数をチェック
                Range targetCell = worksheet.Range["K4"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // MID関数とC4,3,4の組み合わせをチェック
                    return normalizedFormula.Contains("MID(") && 
                           normalizedFormula.Contains("C4,3,4") &&
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
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_5_03_Impl(string filePath)
        {
            // 要件書: 受注一覧シートのL列に、商品IDの右から3文字を表示します。
            // RIGHT関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "受注一覧");
                if (worksheet == null) return false;
                
                // L4セルのRIGHT関数をチェック
                Range targetCell = worksheet.Range["L4"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // RIGHT関数とC4,3の組み合わせをチェック
                    return normalizedFormula.Contains("RIGHT(") && 
                           normalizedFormula.Contains("C4,3") &&
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
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_5_04_Impl(string filePath)
        {
            // 要件書: 受注一覧シートのM列に、小文字の商品IDを大文字で表示します。
            // UPPER関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "受注一覧");
                if (worksheet == null) return false;
                
                // M4セルのUPPER関数をチェック
                Range targetCell = worksheet.Range["M4"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // UPPER関数とC4の組み合わせをチェック
                    return normalizedFormula.Contains("UPPER(") && 
                           normalizedFormula.Contains("C4") &&
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
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_5_05_Impl(string filePath)
        {
            // 要件書: 商品マスタシートのE列に、商品名の文字数を表示します。
            // LEN関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "商品マスタ");
                if (worksheet == null) return false;
                
                // E4セルのLEN関数をチェック
                Range targetCell = worksheet.Range["E4"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // LEN関数とB4の組み合わせをチェック
                    return normalizedFormula.Contains("LEN(") && 
                           normalizedFormula.Contains("B4") &&
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

        private ListObject FindTable(Worksheet worksheet)
        {
            try
            {
                foreach (ListObject table in worksheet.ListObjects)
                {
                    return table; // 最初のテーブルを返す
                }
            }
            catch
            {
                // Error accessing tables
            }
            return null;
        }

    }
}
