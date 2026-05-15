using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group3
{
    public class ExcelChecker3_1
    {
        public bool CheckTask_3_1_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_1_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_3_1_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_1_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_3_1_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_1_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_3_1_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_1_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_3_1_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_3_1_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_3_1_01(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null) return false;
                
                // タスク1-1: 文化祭コピーシートの表にL8:Q16をコピーして貼り付け
                worksheet = FindWorksheet(workbook, "文化祭コピー");
                if (worksheet == null) return false;
                
                // B8:G16にデータがコピーされているかチェック
                Range targetRange = worksheet.Range["B8:G16"];
                Range sourceRange = worksheet.Range["L8:Q16"];
                
                // ソース範囲にデータが存在するかチェック
                bool sourceHasData = false;
                foreach (Range cell in sourceRange.Cells)
                {
                    if (cell.Value2 != null)
                    {
                        sourceHasData = true;
                        break;
                    }
                }
                
                // ターゲット範囲にデータがコピーされているかチェック
                bool targetHasData = false;
                foreach (Range cell in targetRange.Cells)
                {
                    if (cell.Value2 != null)
                    {
                        targetHasData = true;
                        break;
                    }
                }
                
                // ソース範囲にデータがあり、ターゲット範囲にもデータがコピーされている場合のみ成功
                return sourceHasData && targetHasData;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_3_1_02(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null) return false;
                
                // タスク1-2: 文化祭切り取りシートの表にL8:Q16を切り取って貼り付け
                worksheet = FindWorksheet(workbook, "文化祭切り取り");
                if (worksheet == null) return false;
                
                // B8:G16にデータが移動され、L8:Q16が空になっているかチェック
                Range targetRange = worksheet.Range["B8:G16"];
                Range sourceRange = worksheet.Range["L8:Q16"];
                
                // 1. ソース範囲が空になっているかチェック
                bool sourceIsEmpty = true;
                foreach (Range cell in sourceRange.Cells)
                {
                    if (cell.Value2 != null && !string.IsNullOrWhiteSpace(cell.Value2.ToString()))
                    {
                        sourceIsEmpty = false;
                        break;
                    }
                }
                
                // 2. ターゲット範囲にデータが存在するかチェック
                bool hasTargetData = false;
                int targetDataCount = 0;
                foreach (Range cell in targetRange.Cells)
                {
                    if (cell.Value2 != null && !string.IsNullOrWhiteSpace(cell.Value2.ToString()))
                    {
                        hasTargetData = true;
                        targetDataCount++;
                    }
                }
                
                // 切り取り貼り付けの判定：
                // - ソース範囲が空になっている
                // - ターゲット範囲にデータが存在する
                // - ターゲット範囲に適度な量のデータがある（最低3つ以上）
                return sourceIsEmpty && hasTargetData && targetDataCount >= 3;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_3_1_03(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null) return false;
                
                // タスク1-3: 文化祭切り取りシートのテーブルのH列の式をオートフィル
                worksheet = FindWorksheet(workbook, "文化祭切り取り");
                if (worksheet == null) return false;
                
                // H8からH16までの範囲で数式がコピーされているかチェック
                Range h8Cell = worksheet.Range["H8"];
                Range targetRange = worksheet.Range["H9:H16"];
                
                // 1. H8に数式があるかチェック
                bool h8HasFormula = h8Cell.HasFormula is bool && (bool)h8Cell.HasFormula;
                if (!h8HasFormula) return false;
                
                // 2. H8の数式を取得
                string h8Formula = h8Cell.Formula as string;
                if (string.IsNullOrEmpty(h8Formula)) return false;
                
                // 3. H9:H16の各セルに数式があるかチェック
                int formulaCount = 0;
                int validFormulaCount = 0;
                
                foreach (Range cell in targetRange.Cells)
                {
                    bool cellHasFormula = cell.HasFormula is bool && (bool)cell.HasFormula;
                    if (!cellHasFormula) return false;
                    
                    string cellFormula = cell.Formula as string;
                    if (!string.IsNullOrEmpty(cellFormula))
                    {
                        formulaCount++;
                        
                        // 4. 数式が相対参照で適切に調整されているかチェック（より柔軟に）
                        // H8の数式に含まれる数字（行番号）が、対象セルの行番号に調整されているかを確認
                        bool hasRelativeReference = false;
                        
                        // H8の数式から行番号パターンを抽出して比較
                        // 例：H8が =A8*B8 なら、H9は =A9*B9 になっているはず
                        for (int row = 1; row <= 20; row++) // 行1-20をチェック
                        {
                            string rowPattern = row.ToString();
                            if (h8Formula.Contains(rowPattern))
                            {
                                // 対象セルの行番号に調整されているかチェック
                                string expectedPattern = cell.Row.ToString();
                                if (cellFormula.Contains(expectedPattern))
                                {
                                    hasRelativeReference = true;
                                    break;
                                }
                            }
                        }
                        
                        if (hasRelativeReference)
                        {
                            validFormulaCount++;
                        }
                    }
                }
                
                // 5. 最低限、H9:H16の半分以上のセルに数式があり、
                //    そのうちの半分以上が相対参照で適切に調整されていることを確認
                return formulaCount >= (targetRange.Cells.Count / 2) && 
                       validFormulaCount >= (formulaCount / 2);
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_3_1_04(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null) return false;
                
                // タスク1-4: 担当者マスターシートの表にG11:J16をコピーして列幅を揃えて貼り付け
                worksheet = FindWorksheet(workbook, "担当者マスター");
                if (worksheet == null) return false;
                
                // A11:D16にデータがコピーされているかチェック
                Range targetRange = worksheet.Range["A11:D16"];
                
                bool hasData = false;
                foreach (Range cell in targetRange.Cells)
                {
                    if (cell.Value2 != null)
                    {
                        hasData = true;
                        break;
                    }
                }
                
                // データがコピーされていない場合は失敗
                if (!hasData) return false;
                
                // 列幅が揃っているかチェック（A列、B列、C列、D列の幅が同じか）
                Range columnA = (Range)worksheet.Columns["A:A"];
                Range columnB = (Range)worksheet.Columns["B:B"];
                Range columnC = (Range)worksheet.Columns["C:C"];
                Range columnD = (Range)worksheet.Columns["D:D"];
                
                double columnAWidth = Convert.ToDouble(columnA.ColumnWidth);
                double columnBWidth = Convert.ToDouble(columnB.ColumnWidth);
                double columnCWidth = Convert.ToDouble(columnC.ColumnWidth);
                double columnDWidth = Convert.ToDouble(columnD.ColumnWidth);
                
                // 列幅が揃っているかチェック（許容誤差0.1以内）
                bool columnsWidthMatched = Math.Abs(columnAWidth - columnBWidth) < 0.1 &&
                                         Math.Abs(columnBWidth - columnCWidth) < 0.1 &&
                                         Math.Abs(columnCWidth - columnDWidth) < 0.1;
                
                return columnsWidthMatched;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_3_1_05(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                
                if (workbook == null) return false;
                
                // タスク1-5: 担当者マスターシートの売上にL11:L16をリンク貼り付け
                worksheet = FindWorksheet(workbook, "担当者マスター");
                if (worksheet == null) return false;
                
                // E11:E16にリンク数式があるかチェック
                Range targetRange = worksheet.Range["E11:E16"];
                
                foreach (Range cell in targetRange.Cells)
                {
                    bool hasFormula = cell.HasFormula is bool && (bool)cell.HasFormula;
                    if (hasFormula)
                    {
                        string formula = cell.Formula as string;
                        // リンク貼り付けは他のセルへの参照になる（L11:L16の範囲を参照）
                        if (formula != null && (formula.Contains("L11") || formula.Contains("L12") || 
                            formula.Contains("L13") || formula.Contains("L14") || 
                            formula.Contains("L15") || formula.Contains("L16")))
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
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private string GetCurrentExcelFilePath()
        {
            try
            {
                Application excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
            }
            catch
            {
                // Excel is not running or no active workbook
            }
            return string.Empty;
        }

        private Worksheet FindWorksheet(Workbook workbook, string worksheetName)
        {
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                if (worksheet.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }
            return null;
        }
    }
}
