using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_8
    {
        public bool CheckTask_1_8_01() => RunCheck(CheckTask_1_8_01_Impl);
        public bool CheckTask_1_8_02() => RunCheck(CheckTask_1_8_02_Impl);
        public bool CheckTask_1_8_03() => RunCheck(CheckTask_1_8_03_Impl);
        public bool CheckTask_1_8_04() => RunCheck(CheckTask_1_8_04_Impl);
        public bool CheckTask_1_8_05() => RunCheck(CheckTask_1_8_05_Impl);
        public bool CheckTask_1_8_06() => RunCheck(CheckTask_1_8_06_Impl);

        private bool RunCheck(Func<string, bool> task)
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return task(filePath);
            }
            catch
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        // ---------------------------------------------------------
        // Task 9-1: 名前定義「氏名」→ 学生名簿!C5:C24
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
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                bool foundShimei = false;
                string shimeiSheet = null;
                string shimeiAddress = null;
                var wrongNamesOnExpectedRange = new List<string>();

                foreach (Excel.Name name in workbook.Names)
                {
                    try
                    {
                        Excel.Range range = name.RefersToRange;
                        if (range == null)
                            continue;

                        string sheet = range.Worksheet.Name;
                        string address = range.Address.Replace("$", "");
                        bool onExpectedRange = string.Equals(sheet, "学生名簿", StringComparison.OrdinalIgnoreCase)
                            && string.Equals(address, "C5:C24", StringComparison.OrdinalIgnoreCase);

                        string shortName = GetShortDefinedName(name.Name);
                        bool isShimei = shortName == "氏名";

                        if (isShimei)
                        {
                            foundShimei = true;
                            shimeiSheet = sheet;
                            shimeiAddress = address;
                            if (onExpectedRange)
                                return true;
                        }
                        else if (onExpectedRange)
                        {
                            if (!string.IsNullOrEmpty(shortName) && !wrongNamesOnExpectedRange.Contains(shortName))
                                wrongNamesOnExpectedRange.Add(shortName);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (!foundShimei)
                {
                    if (wrongNamesOnExpectedRange.Count > 0)
                    {
                        return Miss(
                            $"名前「氏名」がなく、同じ範囲に名前「{Quote(JoinNames(wrongNamesOnExpectedRange))}」が設定されています。");
                    }
                    return Miss("名前「氏名」が定義されていません。");
                }

                if (!string.Equals(shimeiSheet, "学生名簿", StringComparison.OrdinalIgnoreCase))
                    ExcelScoreExplanation.Note($"名前「氏名」の参照シートが「{Quote(shimeiSheet ?? "（不明）")}」になっています。");
                if (!string.Equals(shimeiAddress, "C5:C24", StringComparison.OrdinalIgnoreCase))
                    ExcelScoreExplanation.Note($"名前「氏名」の参照範囲が「{Quote(shimeiAddress ?? "（不明）")}」になっています。");
                return false;
            }
            catch
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        // ---------------------------------------------------------
        // Task 9-2: 名前「開催日」の値を 5/10 に変更
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
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                foreach (Excel.Name name in workbook.Names)
                {
                    try
                    {
                        if (name.Name != "開催日" && !name.Name.EndsWith("!開催日"))
                            continue;

                        Excel.Range range = name.RefersToRange;
                        if (range == null)
                            return Miss("名前「開催日」の参照先が取得できません。");

                        string value = range.Text != null ? range.Text.ToString() : "";

                        if (!string.IsNullOrEmpty(value) && value.Contains("5/10"))
                            return true;

                        return Miss($"名前「開催日」の値が「{Quote(string.IsNullOrEmpty(value) ? "（空）" : value)}」になっています。");
                    }
                    catch
                    {
                        continue;
                    }
                }

                return Miss("名前「開催日」が見つかりません。");
            }
            catch
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        // ---------------------------------------------------------
        // Task 9-3: 売上報告!J5 = SUM(売上合計)（セル範囲参照なし）
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
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                const string sheet = "売上報告";
                const string cell = "J5";
                worksheet = FindWorksheet(workbook, sheet);
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Excel.Range targetCell = worksheet.Range[cell];
                string formula = targetCell.Formula as string;
                if (string.IsNullOrWhiteSpace(formula))
                    return Miss($"{SheetCell(sheet, cell)}に数式がありません。");

                string compact = CompactFormula(formula);
                bool hasSumName = compact.IndexOf("SUM(売上合計)", StringComparison.OrdinalIgnoreCase) >= 0;
                bool hasColon = compact.Contains(":");

                if (hasSumName && !hasColon)
                    return true;

                if (!hasSumName)
                    ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}が名前付き範囲「売上合計」を使ったSUMになっていません。");
                if (hasColon)
                    ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の数式にセル範囲の参照が含まれています。");
                return false;
            }
            catch
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ---------------------------------------------------------
        // Task 9-4: 学生名簿!G5:G24 = CONCAT(En:Fn) または CONCAT(En,Fn)
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
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                const string sheet = "学生名簿";
                worksheet = FindWorksheet(workbook, sheet);
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Excel.Range g5 = worksheet.Range["G5"];
                string g5Formula = g5.Formula as string;
                if (string.IsNullOrWhiteSpace(g5Formula))
                    return Miss($"{SheetCell(sheet, "G5")}に数式がありません。");

                if (!IsConcatEf(g5Formula, 5))
                {
                    ExplainConcatFail(sheet, "G5", g5Formula, "CONCAT");
                    return false;
                }

                var missingOrWrong = new List<string>();
                foreach (Excel.Range cell in worksheet.Range["G6:G24"].Cells)
                {
                    int row = cell.Row;
                    string formula = cell.Formula as string;
                    if (!IsConcatEf(formula, row))
                        missingOrWrong.Add(cell.Address[false, false]);
                }

                if (missingOrWrong.Count == 0)
                    return true;

                return Miss($"シート「{sheet}」の{JoinNames(missingOrWrong)}にCONCAT関数がありません。");
            }
            catch
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ---------------------------------------------------------
        // Task 9-5: 担当者リスト!H5:H19 = CONCAT(Gn,"@rabbit.ac.jp")
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
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                const string sheet = "担当者リスト";
                worksheet = FindWorksheet(workbook, sheet);
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Excel.Range h5 = worksheet.Range["H5"];
                string h5Formula = h5.Formula as string;
                if (string.IsNullOrWhiteSpace(h5Formula))
                    return Miss($"{SheetCell(sheet, "H5")}に数式がありません。");

                if (!IsConcatMail(h5Formula, 5))
                {
                    ExplainConcatFail(sheet, "H5", h5Formula, "CONCAT");
                    return false;
                }

                var missingOrWrong = new List<string>();
                foreach (Excel.Range cell in worksheet.Range["H6:H19"].Cells)
                {
                    int row = cell.Row;
                    string formula = cell.Formula as string;
                    if (!IsConcatMail(formula, row))
                        missingOrWrong.Add(cell.Address[false, false]);
                }

                if (missingOrWrong.Count == 0)
                    return true;

                return Miss($"シート「{sheet}」の{JoinNames(missingOrWrong)}にCONCAT関数がありません。");
            }
            catch
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ---------------------------------------------------------
        // Task 9-6: 申込一覧!G5:G174 = CONCAT(Bn,"-",En)
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
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                const string sheet = "申込一覧";
                worksheet = FindWorksheet(workbook, sheet);
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Excel.Range g5 = worksheet.Range["G5"];
                string g5Formula = g5.Formula as string;
                if (string.IsNullOrWhiteSpace(g5Formula))
                    return Miss($"{SheetCell(sheet, "G5")}に数式がありません。");

                if (!IsConcatCampus(g5Formula, 5))
                {
                    ExplainConcatFail(sheet, "G5", g5Formula, "CONCAT");
                    return false;
                }

                var missingOrWrong = new List<string>();
                foreach (Excel.Range cell in worksheet.Range["G6:G174"].Cells)
                {
                    int row = cell.Row;
                    string formula = cell.Formula as string;
                    if (!IsConcatCampus(formula, row))
                        missingOrWrong.Add(cell.Address[false, false]);
                }

                if (missingOrWrong.Count == 0)
                    return true;

                return Miss($"シート「{sheet}」の{JoinNames(missingOrWrong)}にCONCAT関数がありません。");
            }
            catch
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // Helpers
        private string GetCurrentExcelFilePath()
        {
            try
            {
                Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                return excelApp.ActiveWorkbook?.FullName;
            }
            catch
            {
                return null;
            }
        }

        private Excel.Worksheet FindWorksheet(Excel.Workbook workbook, string sheetName)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return sheet;
            }
            return null;
        }

        private Excel.Workbook GetWorkbook(Excel.Application excelApp, string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            foreach (Excel.Workbook wb in excelApp.Workbooks)
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

        /// <summary>「学生名簿!氏名」→「氏名」のようにローカル名を取り出す。</summary>
        private static string GetShortDefinedName(string fullName)
        {
            if (string.IsNullOrEmpty(fullName))
                return "";
            int bang = fullName.LastIndexOf('!');
            if (bang >= 0 && bang < fullName.Length - 1)
                return fullName.Substring(bang + 1);
            return fullName;
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

        private static string CompactFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula))
                return "";
            return formula.Replace(" ", "");
        }

        private static bool IsConcatEf(string formula, int row)
        {
            string n = NormalizeFormula(formula ?? "");
            if (n.Contains("$"))
                return false;
            string rangeForm = $"CONCAT(E{row}:F{row})";
            string commaForm = $"CONCAT(E{row},F{row})";
            return n.Contains(rangeForm) || n.Contains(commaForm);
        }

        private static bool IsConcatMail(string formula, int row)
        {
            string n = NormalizeFormula(formula ?? "");
            if (n.Contains("$"))
                return false;
            return n.Contains($"CONCAT(G{row},\"@RABBIT.AC.JP\")")
                || n.Contains($"CONCAT(G{row},'@RABBIT.AC.JP')");
        }

        private static bool IsConcatCampus(string formula, int row)
        {
            string n = NormalizeFormula(formula ?? "");
            if (n.Contains("$"))
                return false;
            return n.Contains($"CONCAT(B{row},\"-\",E{row})")
                || n.Contains($"CONCAT(B{row},'-',E{row})");
        }

        private static void ExplainConcatFail(string sheet, string cell, string formula, string funcName)
        {
            string n = NormalizeFormula(formula ?? "");
            bool hasFunc = n.Contains(funcName + "(");
            if (!hasFunc)
            {
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}に{funcName}関数がありません。");
            }
            else
            {
                ExcelScoreExplanation.Note(
                    $"{SheetCell(sheet, cell)}の引数が「{Quote(DescribeFormulaArgs(formula, funcName))}」になっています。");
            }
            if (n.Contains("$"))
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の数式に「$」が含まれています。");
        }

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
