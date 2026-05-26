using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group3
{
    public class ExcelChecker3_2
    {
        // Public CheckTask methods
        public bool CheckTask_3_2_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_2_01_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_2_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_2_02_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_2_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_2_03_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_2_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_2_04_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        public bool CheckTask_3_2_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_2_05_Impl(filePath);
            }
            catch
            {
                return false;
            }
        }

        // Private implementation methods
        private bool CheckTask_3_2_01_Impl(string filePath)
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
                
                // タスク2-1: キャンパス別試験結果シートの学部学科列にCONCAT関数で学部と学科を結合
                worksheet = FindWorksheet(workbook, "キャンパス別試験結果");
                if (worksheet == null) return false;
                
                // G7セルのCONCAT関数をチェック（E7:F7を結合）
                Range targetCell = worksheet.Range["G7"];
                bool hasFormula = targetCell.HasFormula is bool && (bool)targetCell.HasFormula;
                if (!hasFormula) return false;
                
                string formula = targetCell.Formula as string;
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // CONCAT(E7:F7) または CONCATENATE(E7:F7) をチェック
                    return normalizedFormula.Contains("CONCAT(E7:F7)") || 
                           normalizedFormula.Contains("CONCATENATE(E7:F7)") ||
                           normalizedFormula.Contains("=CONCAT(E7:F7)") ||
                           normalizedFormula.Contains("=CONCATENATE(E7:F7)");
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

        private bool CheckTask_3_2_02_Impl(string filePath)
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
                
                // タスク2-2: キャンパス別試験結果シートの氏名の列C7:C26に左インデントを1つ追加
                worksheet = FindWorksheet(workbook, "キャンパス別試験結果");
                if (worksheet == null) return false;
                
                Range targetRange = worksheet.Range["C7:C26"];
                return CheckIndentLevelDetailed(targetRange, 1);
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

        private bool CheckTask_3_2_03_Impl(string filePath)
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
                
                // タスク2-3: キャンパス別試験結果シートのセルB4内の文字をB4とC4の中央に配置
                worksheet = FindWorksheet(workbook, "キャンパス別試験結果");
                if (worksheet == null) return false;
                
                Range b4Cell = worksheet.Range["B4"];
                // -4108はxlCenter（中央揃え）の定数値
                return CheckAlignmentDetailed(b4Cell, -4108);
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

        private bool CheckTask_3_2_04_Impl(string filePath)
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
                    excelApp.Visible = false;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                // タスク2-4: キャンパス別試験結果シートのB6:C26に「学生」という名前を付ける
                foreach (Name name in workbook.Names)
                {
                    if (name.Name == "学生")
                    {
                        // 参照先がB6:C26かチェック
                        string refersTo = name.RefersTo as string;
                        if (!string.IsNullOrEmpty(refersTo) && 
                            (refersTo.Contains("B6:C26") || refersTo.Contains("$B$6:$C$26")))
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
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private bool CheckTask_3_2_05_Impl(string filePath)
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
                    excelApp.Visible = false;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                // タスク2-5: 「消費税」という名前付き範囲に移動し、数字を「10％」に変更
                foreach (Name name in workbook.Names)
                {
                    if (name.Name == "消費税")
                    {
                        // 参照先の値をチェック
                        try
                        {
                            Range range = name.RefersToRange;
                            if (range != null)
                            {
                                string value = range.Text.ToString();
                                if (!string.IsNullOrEmpty(value) && 
                                    (value.Contains("10%") || value.Contains("10％") || value == "10%" || value == "10％"))
                                {
                                    return true;
                                }
                            }
                        }
                        catch
                        {
                            // 参照先が取得できない場合は継続
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

        // インデントレベルの詳細確認
        private bool CheckIndentLevelDetailed(Range range, int expectedLevel)
        {
            try
            {
                bool allCellsCorrect = true;
                int checkedCells = 0;
                
                foreach (Range cell in range.Cells)
                {
                    checkedCells++;
                    int indentLevel = 0;
                    
                    try
                    {
                        var indentProperty = cell.GetType().GetProperty("IndentLevel");
                        if (indentProperty != null)
                        {
                            var indentValue = indentProperty.GetValue(cell);
                            indentLevel = Convert.ToInt32(indentValue);
                        }
                    }
                    catch
                    {
                        // インデントレベルが取得できない場合は失敗
                        return false;
                    }
                    
                    if (indentLevel != expectedLevel)
                    {
                        allCellsCorrect = false;
                        break;
                    }
                }
                
                return allCellsCorrect && checkedCells > 0;
            }
            catch
            {
                return false;
            }
        }

        // 文字配置の詳細確認
        private bool CheckAlignmentDetailed(Range cell, int expectedAlignment)
        {
            try
            {
                int horizontalAlignment = 0;
                
                try
                {
                    horizontalAlignment = (int)cell.HorizontalAlignment;
                }
                catch
                {
                    var alignmentProperty = cell.GetType().GetProperty("HorizontalAlignment");
                    if (alignmentProperty != null)
                    {
                        var alignmentValue = alignmentProperty.GetValue(cell);
                        horizontalAlignment = (int)alignmentValue;
                    }
                }
                
                return horizontalAlignment == expectedAlignment;
            }
            catch
            {
                return false;
            }
        }
    }
}
