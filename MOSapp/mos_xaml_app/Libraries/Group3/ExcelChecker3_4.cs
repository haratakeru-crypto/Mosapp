using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group3
{
    public class ExcelChecker3_4
    {
        // Public CheckTask methods
        public bool CheckTask_3_4_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_4_01_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_4_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_4_02_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_4_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_4_03_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_4_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_4_04_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_4_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_4_05_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        // Private implementation methods
        private bool CheckTask_3_4_01_Impl(string filePath)
        {
            // 要件書: 売上表シートの表を、テーブルに変換します。テーブルスタイルは「青、テーブルスタイル(中間)2」とします。
            // ExcelChecker1_2のCheckTask_1_2_03_Implを参考に、テーブルスタイルをチェック
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
                
                worksheet = FindWorksheet(workbook, "売上表");
                if (worksheet == null) return false;
                
                // テーブルを検索
                ListObject table = FindTable(worksheet);
                if (table == null) return false;
                
                // テーブルスタイルをチェック
                string tableStyle = GetTableStyleName(table);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Current Style Name: {tableStyle}");
                
                // 「青、テーブルスタイル(中間)2」をチェック
                string[] validStyles = {
                    "TableStyleMedium2",
                    "Medium2",
                    "Medium 2",
                    "テーブルスタイル（中間）2",
                    "2"
                };
                
                bool matchFound = false;
                foreach (string validStyle in validStyles)
                {
                    if (!string.IsNullOrEmpty(tableStyle) && 
                        tableStyle.EndsWith(validStyle, StringComparison.OrdinalIgnoreCase))
                    {
                        matchFound = true;
                        break;
                    }
                }
                
                return matchFound;
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

        private bool CheckTask_3_4_02_Impl(string filePath)
        {
            // 要件書: 売上表シートのテーブルのデータを、商品名の昇順に並べ替えます。
            // ExcelChecker2_10のCheckTask_2_10_02_Implを参考に、並べ替えをチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            ListObject table = null;
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
                
                worksheet = FindWorksheet(workbook, "売上表");
                if (worksheet == null) return false;
                
                table = FindTable(worksheet);
                if (table == null) return false;
                
                // オートフィルターが適用されているかチェック
                if (table.AutoFilter == null) return false;
                
                // 並べ替えが適用されているかチェック（簡易的な判定）
                AutoFilter autoFilter = table.AutoFilter;
                if (autoFilter.Sort != null)
                {
                    return true; // 並べ替えが適用されている
                }
                
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (table != null) Marshal.ReleaseComObject(table);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_4_03_Impl(string filePath)
        {
            // 要件書: 売上表シートのテーブルに、「集計行」を追加します。
            // テーブルの集計行が有効になっているかチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            ListObject table = null;
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
                
                worksheet = FindWorksheet(workbook, "売上表");
                if (worksheet == null) return false;
                
                table = FindTable(worksheet);
                if (table == null) return false;
                
                // 集計行が有効になっているかチェック
                bool showTotalsRow = table.ShowTotals;
                return showTotalsRow;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (table != null) Marshal.ReleaseComObject(table);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_4_04_Impl(string filePath)
        {
            // 要件書: 売上表シートの集計行のセルB22に、個数の合計を表示します。
            // 集計行のセルB22にSUM関数またはSUBTOTAL関数があるかチェック
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
                
                worksheet = FindWorksheet(workbook, "売上表");
                if (worksheet == null) return false;
                
                // B22セルの数式をチェック
                Range targetCell = worksheet.Range["B22"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // SUM関数またはSUBTOTAL関数（合計）をチェック
                    return normalizedFormula.Contains("SUM(") || 
                           normalizedFormula.Contains("SUBTOTAL(109") || // 109は合計
                           normalizedFormula.Contains("SUBTOTAL(9"); // 9も合計
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

        private bool CheckTask_3_4_05_Impl(string filePath)
        {
            // 要件書: 売上表シートの集計行のセルH22に、金額の平均を表示します。
            // 集計行のセルH22にAVERAGE関数またはSUBTOTAL関数（平均）があるかチェック
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
                
                worksheet = FindWorksheet(workbook, "売上表");
                if (worksheet == null) return false;
                
                // H22セルの数式をチェック
                Range targetCell = worksheet.Range["H22"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // AVERAGE関数またはSUBTOTAL関数（平均）をチェック
                    return normalizedFormula.Contains("AVERAGE(") || 
                           normalizedFormula.Contains("SUBTOTAL(101") || // 101は平均
                           normalizedFormula.Contains("SUBTOTAL(1"); // 1も平均
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

        // テーブルスタイル名を取得
        private string GetTableStyleName(ListObject table)
        {
            try
            {
                dynamic style = table.TableStyle;
                return style.Name;
            }
            catch { return ""; }
        }
    }
}
