using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group3
{
    public class ExcelChecker3_6
    {
        // Public CheckTask methods
        public bool CheckTask_3_6_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_6_01_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_6_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_6_02_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_6_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_6_03_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_6_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_6_04_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_6_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_6_05_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }


        // Private implementation methods
        private bool CheckTask_3_6_01_Impl(string filePath)
        {
            // 要件書: 社員名簿シートのH列に、氏名にフリガナを表示します。
            // PHONETIC関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "社員名簿");
                if (worksheet == null) return false;
                
                // H4セルのPHONETIC関数をチェック
                Range targetCell = worksheet.Range["H4"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // PHONETIC関数とB4の組み合わせをチェック
                    return normalizedFormula.Contains("PHONETIC(") && 
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

        private bool CheckTask_3_6_02_Impl(string filePath)
        {
            // 要件書: 請求書シートのセルC14に、今日の日付を表示します。
            // TODAY関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "請求書");
                if (worksheet == null) return false;
                
                // C14セルのTODAY関数をチェック
                Range targetCell = worksheet.Range["C14"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // TODAY関数をチェック
                    return normalizedFormula.Contains("TODAY()");
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

        private bool CheckTask_3_6_03_Impl(string filePath)
        {
            // 要件書: 請求書シートのセルC15に、現在の日付と時刻を表示します。
            // NOW関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "請求書");
                if (worksheet == null) return false;
                
                // C15セルのNOW関数をチェック
                Range targetCell = worksheet.Range["C15"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // NOW関数をチェック
                    return normalizedFormula.Contains("NOW()");
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

        private bool CheckTask_3_6_04_Impl(string filePath)
        {
            // 要件書: 見積書シートのセルH12に、セルH10とH11の合計を表示します。ただし、H10またはH11が空欄の場合は空欄を表示します。
            // IF関数とOR関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "見積書");
                if (worksheet == null) return false;
                
                // H12セルのIF関数をチェック
                Range targetCell = worksheet.Range["H12"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // IF関数、OR関数、H10=""またはH11=""の条件、H10+H11の合計、空文字列をチェック
                    bool hasIF = normalizedFormula.Contains("IF(");
                    bool hasOR = normalizedFormula.Contains("OR(");
                    bool hasCondition = (normalizedFormula.Contains("H10=\"\"") || normalizedFormula.Contains("H11=\"\""));
                    bool hasSum = normalizedFormula.Contains("H10+H11");
                    bool hasEmptyString = normalizedFormula.Contains("\"\"");
                    
                    return hasIF && hasOR && hasCondition && hasSum && hasEmptyString;
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

        private bool CheckTask_3_6_05_Impl(string filePath)
        {
            // 要件書: 見積書シートのセルH18に、セルF18とG18の積を表示します。エラーが表示される場合は、0を表示します。
            // IFERROR関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "見積書");
                if (worksheet == null) return false;
                
                // H18セルのIFERROR関数をチェック
                Range targetCell = worksheet.Range["H18"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // IFERROR関数、F18*G18の積、0をチェック
                    bool hasIFERROR = normalizedFormula.Contains("IFERROR(");
                    bool hasProduct = normalizedFormula.Contains("F18*G18");
                    bool hasZero = normalizedFormula.Contains(",0");
                    
                    return hasIFERROR && hasProduct && hasZero;
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
