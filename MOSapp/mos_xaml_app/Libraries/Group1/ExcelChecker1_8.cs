using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Libraries.Group1
{
    public class ExcelChecker1_8
    {
        private static readonly string TARGET_FILE_PATH =
            MOSExcelMogiApp.Infrastructure.DataPathHelper.GetWorkingFilePath(1, 8);

        // Public wrappers
        public bool CheckTask_1_8_01() => RunCheck(CheckTask_1_8_01_Impl);
        public bool CheckTask_1_8_02() => RunCheck(CheckTask_1_8_02_Impl);
        public bool CheckTask_1_8_03() => RunCheck(CheckTask_1_8_03_Impl);
        public bool CheckTask_1_8_04() => RunCheck(CheckTask_1_8_04_Impl);
        public bool CheckTask_1_8_05() => RunCheck(CheckTask_1_8_05_Impl);
        public bool CheckTask_1_8_06() => RunCheck(CheckTask_1_8_06_Impl);

        // Common Runner
        private bool RunCheck(Func<string, bool> task)
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath)) return false;
                return task(filePath);
            }
            catch { return false; }
        }

        // ---------------------------------------------------------
        // Task 8-1: 名前定義「氏名」
        // ---------------------------------------------------------
        private bool CheckTask_1_8_01_Impl(string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            try
            {
                try { excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Excel.Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                // 名前「氏名」を探す（スコープの違いを考慮して EndsWith を使用）
                foreach (Excel.Name name in workbook.Names)
                {
                    try
                    {
                        // "氏名" または "学生名簿!氏名" などにマッチさせる
                        if (name.Name == "氏名" || name.Name.EndsWith("!氏名"))
                        {
                            Excel.Range range = name.RefersToRange;
                            if (range != null)
                            {
                                // 参照先シートが「学生名簿」であること
                                if (range.Worksheet.Name == "学生名簿")
                                {
                                    // 参照先アドレスが C5:C24 であること
                                    // Address は $C$5:$C$24 のように返るので $ を削除して比較
                                    string address = range.Address.Replace("$", "");
                                    if (address == "C5:C24")
                                    {
                                        return true;
                                    }
                                }
                            }
                        }
                    }
                    catch { continue; }
                }
                return false;
            }
            catch { return false; }
        }

        // ---------------------------------------------------------
        // Task 8-2: 名前範囲変更「開催日」
        // ---------------------------------------------------------
        private bool CheckTask_1_8_02_Impl(string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            try
            {
                try { excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Excel.Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                foreach (Excel.Name name in workbook.Names)
                {
                    try
                    {
                        if (name.Name == "開催日" || name.Name.EndsWith("!開催日"))
                        {
                            Excel.Range range = name.RefersToRange;
                            if (range != null)
                            {
                                string value = range.Text.ToString();
                                if (!string.IsNullOrEmpty(value) && value.Contains("5/10"))
                                {
                                    return true;
                                }
                            }
                        }
                    }
                    catch { continue; }
                }
                return false;
            }
            catch { return false; }
        }

        // ---------------------------------------------------------
        // Task 8-3: SUM関数 (名前「売上合計」を使用)
        // ---------------------------------------------------------
        private bool CheckTask_1_8_03_Impl(string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                try { excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Excel.Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "売上報告");
                if (worksheet == null) return false;
                
                Excel.Range targetCell = worksheet.Range["J5"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalized = formula.Replace(" ", "").ToUpper();

                    // 1. SUM(売上合計) が含まれているか（必須）
                    bool hasName = normalized.Contains("SUM(売上合計)") || normalized.Contains("=SUM(売上合計)");
                    
                    if (!hasName) return false;

                    // 2. セル番地（数字）が含まれていないかチェック（念のため）
                    // 名前定義「売上合計」を使っていれば、数式中に "5" や "24" といった行番号は出ないはず
                    // ただし、シート名などに数字が含まれる可能性を考慮し、慎重に判定
                    // ここでは簡易的に「:」が含まれていたら範囲指定（C5:C24）が残っているとみなす
                    if (normalized.Contains(":"))
                    {
                        // SUM(売上合計) があるのに : があるのはおかしい（SUM(売上合計, C5:C24)などの場合）
                        return false;
                    }

                    return true;
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ---------------------------------------------------------
        // Task 8-4: CONCAT関数 (氏名結合)
        // ---------------------------------------------------------
        private bool CheckTask_1_8_04_Impl(string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                try { excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Excel.Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "学生名簿");
                if (worksheet == null) return false;
                
                Excel.Range targetCell = worksheet.Range["G5"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalized = formula.Replace(" ", "").ToUpper();
                    return normalized.Contains("CONCAT(E5:F5)") || normalized.Contains("=CONCAT(E5:F5)")
                        || normalized.Contains("CONCAT(E5,F5)") || normalized.Contains("=CONCAT(E5,F5)");
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ---------------------------------------------------------
        // Task 8-5: CONCAT関数 (メール)
        // ---------------------------------------------------------
        private bool CheckTask_1_8_05_Impl(string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                try { excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Excel.Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "担当者リスト");
                if (worksheet == null) return false;
                
                Excel.Range targetCell = worksheet.Range["H5"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalized = formula.Replace(" ", "").ToUpper();
                    // 教材によってドメインが @RABBIT.AC.JP または @WIN.JP の場合があるため両方許容
                    bool rabbit = normalized.Contains("CONCAT(G5,\"@RABBIT.AC.JP\")") || normalized.Contains("=CONCAT(G5,\"@RABBIT.AC.JP\")");
                    bool win = normalized.Contains("CONCAT(G5,\"@WIN.JP\")") || normalized.Contains("=CONCAT(G5,\"@WIN.JP\")");
                    return rabbit || win;
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ---------------------------------------------------------
        // Task 8-6: CONCAT関数 (所属)
        // ---------------------------------------------------------
        private bool CheckTask_1_8_06_Impl(string filePath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                try { excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Excel.Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "申込一覧");
                if (worksheet == null) return false;
                
                Excel.Range targetCell = worksheet.Range["G5"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalized = formula.Replace(" ", "").ToUpper();
                    return normalized.Contains("CONCAT(B5,\"-\",E5)") || normalized.Contains("=CONCAT(B5,\"-\",E5)");
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // Helpers
        private string GetCurrentExcelFilePath()
        {
            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
                return null;
            }
            catch { return null; }
        }

        private Excel.Worksheet FindWorksheet(Excel.Workbook workbook, string sheetName)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase)) return sheet;
            }
            return null;
        }

        private Excel.Workbook GetWorkbook(Excel.Application excelApp, string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            foreach (Excel.Workbook wb in excelApp.Workbooks)
            {
                if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                    wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) return wb;
            }
            return null;
        }
    }
}
