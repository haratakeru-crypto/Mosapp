using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_7
    {
        private static readonly string TARGET_FILE_PATH =
            MOSExcelMogiApp.Infrastructure.DataPathHelper.GetWorkingFilePath(1, 7);

        public bool CheckTask_1_7_01() => RunCheck(CheckTask_1_7_01_Impl);
        public bool CheckTask_1_7_02() => RunCheck(CheckTask_1_7_02_Impl);
        public bool CheckTask_1_7_03() => RunCheck(CheckTask_1_7_03_Impl);
        public bool CheckTask_1_7_04() => RunCheck(CheckTask_1_7_04_Impl);
        public bool CheckTask_1_7_05() => RunCheck(CheckTask_1_7_05_Impl);
        public bool CheckTask_1_7_06() => RunCheck(CheckTask_1_7_06_Impl);
        public bool CheckTask_1_7_07() => RunCheck(CheckTask_1_7_07_Impl);

        private bool RunCheck(Func<string, bool> task)
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return task(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        public string ValidateProject_7()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                    return "警告: project7.xlsxは既に開いています。ファイルを閉じてから再実行してください。";

                if (!File.Exists(TARGET_FILE_PATH))
                    return "エラー: project7.xlsxが見つかりません。";

                var results = new List<string>();
                results.Add($"Task 7-1 (数式コピー): {(CheckTask_1_7_01_Impl(TARGET_FILE_PATH) ? "OK" : "NG")}");
                results.Add($"Task 7-2 (MAX関数): {(CheckTask_1_7_02_Impl(TARGET_FILE_PATH) ? "OK" : "NG")}");
                results.Add($"Task 7-3 (COUNT関数): {(CheckTask_1_7_03_Impl(TARGET_FILE_PATH) ? "OK" : "NG")}");
                results.Add($"Task 7-4 (COUNTBLANK関数): {(CheckTask_1_7_04_Impl(TARGET_FILE_PATH) ? "OK" : "NG")}");
                results.Add($"Task 7-5 (RANDBETWEEN関数): {(CheckTask_1_7_05_Impl(TARGET_FILE_PATH) ? "OK" : "NG")}");
                results.Add($"Task 7-6 (LEFT関数): {(CheckTask_1_7_06_Impl(TARGET_FILE_PATH) ? "OK" : "NG")}");
                results.Add($"Task 7-7 (UNIQUE関数): {(CheckTask_1_7_07_Impl(TARGET_FILE_PATH) ? "OK" : "NG")}");
                return string.Join("\n", results);
            }
            catch (Exception ex)
            {
                return $"エラー: {ex.Message}";
            }
        }

        // ==========================================
        // タスク8-1: H5の数式を H5:H16 までコピー
        // ==========================================
        private bool CheckTask_1_7_01_Impl(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                worksheet = FindWorksheet(workbook, "イベント売上");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Range h5Cell = worksheet.Range["H5"];
                const string sheet = "イベント売上";
                bool h5HasFormula = h5Cell.HasFormula is bool && (bool)h5Cell.HasFormula;
                if (!h5HasFormula)
                    return Miss($"{SheetCell(sheet, "H5")}に数式がありません。");

                var missing = new List<string>();
                Range targetRange = worksheet.Range["H6:H16"];
                foreach (Range cell in targetRange.Cells)
                {
                    bool cellHasFormula = cell.HasFormula is bool && (bool)cell.HasFormula;
                    if (!cellHasFormula)
                    {
                        string addr = cell.Address[false, false];
                        missing.Add(addr);
                    }
                }

                if (missing.Count == 0)
                    return true;

                return Miss($"シート「{sheet}」の{JoinNames(missing)}に数式がありません。");
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        // ==========================================
        // タスク8-2: J3 = MAX(D5:G16)
        // ==========================================
        private bool CheckTask_1_7_02_Impl(string filePath)
        {
            return CheckFormulaTask(
                filePath,
                "イベント売上",
                "J3",
                formula =>
                {
                    string n = NormalizeFormula(formula);
                    bool hasMax = n.Contains("MAX(");
                    bool hasRange = n.Contains("D5:G16")
                        || n.Contains("売上一覧")
                        || n.Contains("[[1日目]:[4日目]]");
                    bool hasAbs = n.Contains("$");
                    return hasMax && hasRange && !hasAbs;
                },
                (formula, n) =>
                {
                    const string sheet = "イベント売上";
                    const string cell = "J3";
                    if (string.IsNullOrWhiteSpace(formula))
                        return Miss($"{SheetCell(sheet, cell)}に数式がありません。");

                    bool hasMax = n.Contains("MAX(");
                    bool hasRange = n.Contains("D5:G16")
                        || n.Contains("売上一覧")
                        || n.Contains("[[1日目]:[4日目]]");
                    bool hasAbs = n.Contains("$");

                    if (!hasMax)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}にMAX関数がありません。");
                    if (hasMax && !hasRange)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の参照範囲が「{Quote(DescribeFormulaArgs(formula, "MAX"))}」になっています。");
                    if (hasAbs)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の数式に絶対参照が含まれています。");
                    if (hasMax && hasRange && !hasAbs)
                        return Miss($"{SheetCell(sheet, cell)}の数式が「{Quote(formula)}」になっています。");
                    return false;
                });
        }

        // ==========================================
        // タスク8-3: I4 = COUNT(L7:L56)
        // ==========================================
        private bool CheckTask_1_7_03_Impl(string filePath)
        {
            return CheckFormulaTask(
                filePath,
                "試験結果",
                "I4",
                formula =>
                {
                    string n = NormalizeFormula(formula);
                    bool hasCount = n.Contains("COUNT(") && !n.Contains("COUNTBLANK(") && !n.Contains("COUNTA(");
                    bool hasRange = n.Contains("L7:L56")
                        || n.Contains("試験結果")
                        || n.Contains("[[合計点]:[合計点]]");
                    bool hasAbs = n.Contains("$");
                    return hasCount && hasRange && !hasAbs;
                },
                (formula, n) =>
                {
                    const string sheet = "試験結果";
                    const string cell = "I4";
                    if (string.IsNullOrWhiteSpace(formula))
                        return Miss($"{SheetCell(sheet, cell)}に数式がありません。");

                    bool hasCount = n.Contains("COUNT(") && !n.Contains("COUNTBLANK(") && !n.Contains("COUNTA(");
                    bool hasRange = n.Contains("L7:L56")
                        || n.Contains("試験結果")
                        || n.Contains("[[合計点]:[合計点]]");
                    bool hasAbs = n.Contains("$");

                    if (!hasCount)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}にCOUNT関数がありません。");
                    if (hasCount && !hasRange)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の参照範囲が「{Quote(DescribeFormulaArgs(formula, "COUNT"))}」になっています。");
                    if (hasAbs)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の数式に絶対参照が含まれています。");
                    if (hasCount && hasRange && !hasAbs)
                        return Miss($"{SheetCell(sheet, cell)}の数式が「{Quote(formula)}」になっています。");
                    return false;
                });
        }

        // ==========================================
        // タスク8-4: J4 = COUNTBLANK(L7:L56)
        // ==========================================
        private bool CheckTask_1_7_04_Impl(string filePath)
        {
            return CheckFormulaTask(
                filePath,
                "試験結果",
                "J4",
                formula =>
                {
                    string n = NormalizeFormula(formula);
                    bool hasFunc = n.Contains("COUNTBLANK(");
                    bool hasRange = n.Contains("L7:L56")
                        || n.Contains("試験結果")
                        || n.Contains("[[合計点]:[合計点]]");
                    bool hasAbs = n.Contains("$");
                    return hasFunc && hasRange && !hasAbs;
                },
                (formula, n) =>
                {
                    const string sheet = "試験結果";
                    const string cell = "J4";
                    if (string.IsNullOrWhiteSpace(formula))
                        return Miss($"{SheetCell(sheet, cell)}に数式がありません。");

                    bool hasFunc = n.Contains("COUNTBLANK(");
                    bool hasRange = n.Contains("L7:L56")
                        || n.Contains("試験結果")
                        || n.Contains("[[合計点]:[合計点]]");
                    bool hasAbs = n.Contains("$");

                    if (!hasFunc)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}にCOUNTBLANK関数がありません。");
                    if (hasFunc && !hasRange)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の参照範囲が「{Quote(DescribeFormulaArgs(formula, "COUNTBLANK"))}」になっています。");
                    if (hasAbs)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の数式に絶対参照が含まれています。");
                    if (hasFunc && hasRange && !hasAbs)
                        return Miss($"{SheetCell(sheet, cell)}の数式が「{Quote(formula)}」になっています。");
                    return false;
                });
        }

        // ==========================================
        // タスク8-5: B7:B56 = RANDBETWEEN(1,8)
        // ==========================================
        private bool CheckTask_1_7_05_Impl(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                const string sheet = "試験結果";
                worksheet = FindWorksheet(workbook, sheet);
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Range b7 = worksheet.Range["B7"];
                string b7Formula = b7.Formula as string;
                if (string.IsNullOrWhiteSpace(b7Formula))
                    return Miss($"{SheetCell(sheet, "B7")}に数式がありません。");

                string b7Norm = NormalizeFormula(b7Formula);
                bool b7HasFunc = b7Norm.Contains("RANDBETWEEN(");
                bool b7HasArgs = b7Norm.Contains("1,8");
                if (!b7HasFunc || !b7HasArgs)
                {
                    if (!b7HasFunc)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, "B7")}にRANDBETWEEN関数がありません。");
                    if (b7HasFunc && !b7HasArgs)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, "B7")}の引数が「{Quote(DescribeFormulaArgs(b7Formula, "RANDBETWEEN"))}」になっています。");
                    return false;
                }

                var missingOrWrong = new List<string>();
                Range fillRange = worksheet.Range["B8:B56"];
                foreach (Range cell in fillRange.Cells)
                {
                    string formula = cell.Formula as string;
                    string n = NormalizeFormula(formula ?? "");
                    if (!(n.Contains("RANDBETWEEN(") && n.Contains("1,8")))
                        missingOrWrong.Add(cell.Address[false, false]);
                }

                if (missingOrWrong.Count == 0)
                    return true;

                return Miss($"シート「{sheet}」の{JoinNames(missingOrWrong)}にRANDBETWEEN(1,8)がありません。");
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        // ==========================================
        // タスク8-6: G7:G56 = LEFT(Cn,2)
        // ==========================================
        private bool CheckTask_1_7_06_Impl(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                const string sheet = "試験結果";
                worksheet = FindWorksheet(workbook, sheet);
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Range g7 = worksheet.Range["G7"];
                string g7Formula = g7.Formula as string;
                if (string.IsNullOrWhiteSpace(g7Formula))
                    return Miss($"{SheetCell(sheet, "G7")}に数式がありません。");

                string g7Norm = NormalizeFormula(g7Formula);
                bool g7HasFunc = g7Norm.Contains("LEFT(");
                bool g7HasArgs = g7Norm.Contains("C7,2");
                bool g7HasAbs = g7Norm.Contains("$");
                if (!g7HasFunc || !g7HasArgs || g7HasAbs)
                {
                    if (!g7HasFunc)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, "G7")}にLEFT関数がありません。");
                    if (g7HasFunc && !g7HasArgs)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, "G7")}の引数が「{Quote(DescribeFormulaArgs(g7Formula, "LEFT"))}」になっています。");
                    if (g7HasAbs)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, "G7")}の数式に絶対参照が含まれています。");
                    return false;
                }

                var missingOrWrong = new List<string>();
                Range fillRange = worksheet.Range["G8:G56"];
                foreach (Range cell in fillRange.Cells)
                {
                    int row = cell.Row;
                    string formula = cell.Formula as string;
                    string n = NormalizeFormula(formula ?? "");
                    string expectedArgs = $"C{row},2";
                    if (!(n.Contains("LEFT(") && n.Contains(expectedArgs) && !n.Contains("$")))
                        missingOrWrong.Add(cell.Address[false, false]);
                }

                if (missingOrWrong.Count == 0)
                    return true;

                return Miss($"シート「{sheet}」の{JoinNames(missingOrWrong)}にLEFT関数（相対参照）がありません。");
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        // ==========================================
        // タスク8-7: I5 = UNIQUE(E5:E174)
        // ==========================================
        private bool CheckTask_1_7_07_Impl(string filePath)
        {
            return CheckFormulaTask(
                filePath,
                "申込一覧",
                "I5",
                formula =>
                {
                    string n = NormalizeFormula(formula);
                    bool hasFunc = n.Contains("UNIQUE(");
                    bool hasRange = n.Contains("E5:E174")
                        || n.Contains("申込一覧")
                        || n.Contains("[[学部]:[学部]]");
                    bool hasAbs = n.Contains("$");
                    return hasFunc && hasRange && !hasAbs;
                },
                (formula, n) =>
                {
                    const string sheet = "申込一覧";
                    const string cell = "I5";
                    if (string.IsNullOrWhiteSpace(formula))
                        return Miss($"{SheetCell(sheet, cell)}に数式がありません。");

                    bool hasFunc = n.Contains("UNIQUE(");
                    bool hasRange = n.Contains("E5:E174")
                        || n.Contains("申込一覧")
                        || n.Contains("[[学部]:[学部]]");
                    bool hasAbs = n.Contains("$");

                    if (!hasFunc)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}にUNIQUE関数がありません。");
                    if (hasFunc && !hasRange)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の参照範囲が「{Quote(DescribeFormulaArgs(formula, "UNIQUE"))}」になっています。");
                    if (hasAbs)
                        ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の数式に絶対参照が含まれています。");
                    if (hasFunc && hasRange && !hasAbs)
                        return Miss($"{SheetCell(sheet, cell)}の数式が「{Quote(formula)}」になっています。");
                    return false;
                });
        }

        private bool CheckFormulaTask(
            string filePath,
            string sheetName,
            string cellAddress,
            Func<string, bool> isPass,
            Func<string, string, bool> explainFail)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                worksheet = FindWorksheet(workbook, sheetName);
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Range targetCell = worksheet.Range[cellAddress];
                string formula = targetCell.Formula as string;
                if (formula != null && isPass(formula))
                    return true;

                string normalized = NormalizeFormula(formula ?? "");
                return explainFail(formula, normalized);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
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
                        return true;
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
                    return excelApp.ActiveWorkbook.FullName;
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
                    return sheet;
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
                    return wb;
            }
            return null;
        }

        private static bool Miss(string reason)
        {
            ExcelScoreExplanation.Note(reason);
            return false;
        }

        private static string SheetCell(string sheetName, string cellAddress)
        {
            return $"シート「{sheetName}」の{cellAddress}";
        }

        private static string Quote(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "（空）";
            string text = value.Replace("\r", "").Replace("\n", " ");
            const int maxLen = 40;
            if (text.Length <= maxLen)
                return text;
            return text.Substring(0, maxLen) + "…";
        }

        private static string JoinNames(IList<string> names)
        {
            if (names == null || names.Count == 0)
                return "";
            const int maxItems = 5;
            if (names.Count <= maxItems)
                return string.Join("、", names);
            return string.Join("、", names.Take(maxItems)) + "ほか";
        }

        private static string NormalizeFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula))
                return "";
            return formula.Replace(" ", "").ToUpperInvariant();
        }

        /// <summary>関数の括弧内をざっくり取り出して理由表示用にする。</summary>
        private static string DescribeFormulaArgs(string formula, string functionName)
        {
            if (string.IsNullOrEmpty(formula) || string.IsNullOrEmpty(functionName))
                return "（不明）";
            string upper = formula.ToUpperInvariant();
            string key = functionName.ToUpperInvariant() + "(";
            int start = upper.IndexOf(key, StringComparison.Ordinal);
            if (start < 0)
                return formula.Trim();
            start += key.Length;
            int depth = 1;
            int i = start;
            for (; i < formula.Length; i++)
            {
                char c = formula[i];
                if (c == '(') depth++;
                else if (c == ')')
                {
                    depth--;
                    if (depth == 0)
                        break;
                }
            }
            if (depth != 0 || i <= start)
                return formula.Trim();
            return formula.Substring(start, i - start).Trim();
        }
    }
}
