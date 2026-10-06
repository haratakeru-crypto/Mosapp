using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_10
    {
        public bool CheckTask_1_10_01() => RunCheck(CheckTask_1_10_01_Impl);
        public bool CheckTask_1_10_02() => RunCheck(CheckTask_1_10_02_Impl);
        public bool CheckTask_1_10_03() => RunCheck(CheckTask_1_10_03_Impl);
        public bool CheckTask_1_10_04() => RunCheck(CheckTask_1_10_04_Impl);
        public bool CheckTask_1_10_05() => RunCheck(CheckTask_1_10_05_Impl);
        public bool CheckTask_1_10_06() => RunCheck(CheckTask_1_10_06_Impl);
        public bool CheckTask_1_10_07() => RunCheck(CheckTask_1_10_07_Impl);
        public bool CheckTask_1_10_08() => RunCheck(CheckTask_1_10_08_Impl);

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
        // 10-1: 担当者リスト!G5:G26 = IF(Fn>5,"あり","なし")
        // ---------------------------------------------------------
        private bool CheckTask_1_10_01_Impl(string filePath)
        {
            return CheckFilledFormula(
                filePath,
                "担当者リスト",
                "G5",
                "G6:G26",
                (formula, row) => IsIfAriNashi(formula, row),
                (sheet, cell, formula) => ExplainIfAriNashi(sheet, cell, formula));
        }

        // ---------------------------------------------------------
        // 10-2: 出張精算!G5:G9 = IF(En>=300,10000,5000)
        // ---------------------------------------------------------
        private bool CheckTask_1_10_02_Impl(string filePath)
        {
            return CheckFilledFormula(
                filePath,
                "出張精算",
                "G5",
                "G6:G9",
                (formula, row) => IsIfKyori(formula, row),
                (sheet, cell, formula) => ExplainIfKyori(sheet, cell, formula));
        }

        // ---------------------------------------------------------
        // 10-3: 売上一覧!G4:G99 = IF(在庫<=13%,"在庫を補充","")
        // ---------------------------------------------------------
        private bool CheckTask_1_10_03_Impl(string filePath)
        {
            return CheckFilledFormula(
                filePath,
                "売上一覧",
                "G4",
                "G5:G99",
                (formula, row) => IsIfStock(formula, row),
                (sheet, cell, formula) => ExplainIfStock(sheet, cell, formula));
        }

        // ---------------------------------------------------------
        // 10-4: 担当者リスト!A5 = SEQUENCE(22,1,1,1)
        // ---------------------------------------------------------
        private bool CheckTask_1_10_04_Impl(string filePath)
        {
            Workbook workbook;
            Worksheet worksheet;
            if (!TryOpenSheet(filePath, "担当者リスト", out workbook, out worksheet))
                return Miss(ExcelScoreExplanation.UnavailableText);

            try
            {
                const string sheet = "担当者リスト";
                const string cell = "A5";
                string formula = worksheet.Range[cell].Formula as string;
                if (string.IsNullOrWhiteSpace(formula))
                    return Miss($"{SheetCell(sheet, cell)}に数式がありません。");

                string n = NormalizeFormula(formula);
                bool hasSeq = n.Contains("SEQUENCE(");
                bool hasArgs = n.Contains("22,1,1,1") || (n.Contains("22") && n.Contains("1,1,1"));
                if (hasSeq && hasArgs)
                    return true;

                if (!hasSeq)
                    ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}にSEQUENCE関数がありません。");
                if (hasSeq && !hasArgs)
                    ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の引数が「{Quote(DescribeFormulaArgs(formula, "SEQUENCE"))}」になっています。");
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ---------------------------------------------------------
        // 10-5: 業務予定!C4 = SEQUENCE(...,0.5)
        // ---------------------------------------------------------
        private bool CheckTask_1_10_05_Impl(string filePath)
        {
            Workbook workbook;
            Worksheet worksheet;
            if (!TryOpenSheet(filePath, "業務予定", out workbook, out worksheet))
                return Miss(ExcelScoreExplanation.UnavailableText);

            try
            {
                const string sheet = "業務予定";
                const string cell = "C4";
                string formula = worksheet.Range[cell].Formula as string;
                if (string.IsNullOrWhiteSpace(formula))
                    return Miss($"{SheetCell(sheet, cell)}に数式がありません。");

                string n = NormalizeFormula(formula);
                bool hasSeq = n.Contains("SEQUENCE(");
                bool hasStep = n.Contains("0.5");
                if (hasSeq && hasStep)
                    return true;

                if (!hasSeq)
                    ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}にSEQUENCE関数がありません。");
                if (hasSeq && !hasStep)
                    ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の目盛りが「{Quote(DescribeFormulaArgs(formula, "SEQUENCE"))}」になっています。");
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ---------------------------------------------------------
        // 10-6: 売上集計/営業予定!D6 = SORT(A6:B14,2,-1)
        // ---------------------------------------------------------
        private bool CheckTask_1_10_06_Impl(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                if (!TryGetWorkbook(filePath, out excelApp, out workbook))
                    return Miss(ExcelScoreExplanation.UnavailableText);

                worksheet = FindWorksheet(workbook, "売上集計") ?? FindWorksheet(workbook, "営業予定");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                string sheet = worksheet.Name;
                const string cell = "D6";
                string formula = worksheet.Range[cell].Formula as string;
                if (string.IsNullOrWhiteSpace(formula))
                    return Miss($"{SheetCell(sheet, cell)}に数式がありません。");

                string n = NormalizeFormula(formula);
                bool hasSort = n.Contains("SORT(");
                bool hasRange = n.Contains("A6:B14");
                bool hasIndex = n.Contains(",2,") || n.Contains(",2)");
                bool hasDesc = n.Contains("-1");
                if (hasSort && hasRange && hasIndex && hasDesc)
                    return true;

                if (!hasSort)
                    ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}にSORT関数がありません。");
                if (hasSort && !hasRange)
                    ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の参照範囲が「{Quote(DescribeFormulaArgs(formula, "SORT"))}」になっています。");
                if (hasSort && hasRange && (!hasIndex || !hasDesc))
                    ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の並べ替え条件が「{Quote(DescribeFormulaArgs(formula, "SORT"))}」になっています。");
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
        // 10-7: 売上一覧!I4:I99 = En*$K$4（K4は絶対参照）
        // ---------------------------------------------------------
        private bool CheckTask_1_10_07_Impl(string filePath)
        {
            Workbook workbook;
            Worksheet worksheet;
            if (!TryOpenSheet(filePath, "売上一覧", out workbook, out worksheet))
                return Miss(ExcelScoreExplanation.UnavailableText);

            try
            {
                const string sheet = "売上一覧";
                string formulaI4 = worksheet.Range["I4"].Formula as string;
                if (string.IsNullOrWhiteSpace(formulaI4))
                    return Miss($"{SheetCell(sheet, "I4")}に数式がありません。");

                if (!IsTaxPrice(formulaI4, 4))
                {
                    ExplainTaxPrice(sheet, "I4", formulaI4, 4);
                    return false;
                }

                var missingOrWrong = new List<string>();
                foreach (Range cell in worksheet.Range["I5:I99"].Cells)
                {
                    int row = cell.Row;
                    string formula = cell.Formula as string;
                    if (!IsTaxPrice(formula, row))
                        missingOrWrong.Add(cell.Address[false, false]);
                }

                if (missingOrWrong.Count == 0)
                    return true;

                return Miss($"シート「{sheet}」の{JoinNames(missingOrWrong)}に税込価格の数式がありません。");
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ---------------------------------------------------------
        // 10-8: 在庫管理!B4 にテキスト取込
        // ---------------------------------------------------------
        private bool CheckTask_1_10_08_Impl(string filePath)
        {
            Workbook workbook;
            Worksheet worksheet;
            if (!TryOpenSheet(filePath, "在庫管理", out workbook, out worksheet))
                return Miss(ExcelScoreExplanation.UnavailableText);

            try
            {
                const string sheet = "在庫管理";
                bool foundImport = false;
                bool foundAtB4 = false;
                string sampleDest = null;

                if (worksheet.QueryTables.Count > 0)
                {
                    foreach (QueryTable qt in worksheet.QueryTables)
                    {
                        Range destination = null;
                        try { destination = qt.Destination; } catch { }
                        if (destination == null)
                            continue;

                        foundImport = true;
                        sampleDest = destination.Address.Replace("$", "");
                        if (destination.Row == 4 && destination.Column == 2)
                        {
                            foundAtB4 = true;
                            break;
                        }
                    }
                }

                if (!foundAtB4 && worksheet.ListObjects.Count > 0)
                {
                    foreach (ListObject lo in worksheet.ListObjects)
                    {
                        Range headerRange = null;
                        try { headerRange = lo.HeaderRowRange; } catch { }
                        if (headerRange == null)
                            continue;

                        // 行位置に関わらず取込ありとみなし、B4 かどうかで正誤を分ける
                        foundImport = true;
                        sampleDest = headerRange.Address.Replace("$", "");
                        if (headerRange.Row == 4 && headerRange.Column == 2)
                        {
                            foundAtB4 = true;
                            break;
                        }
                    }
                }

                if (foundAtB4)
                    return true;

                if (!foundImport)
                    return Miss($"シート「{sheet}」にテキストファイルの取り込みがありません。");

                return Miss($"シート「{sheet}」の取り込み位置が「{Quote(sampleDest ?? "（不明）")}」になっています。");
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ===== fill-range helper =====
        private bool CheckFilledFormula(
            string filePath,
            string sheetName,
            string startCell,
            string fillRangeAddress,
            Func<string, int, bool> isOk,
            Action<string, string, string> explainStart)
        {
            Workbook workbook;
            Worksheet worksheet;
            if (!TryOpenSheet(filePath, sheetName, out workbook, out worksheet))
                return Miss(ExcelScoreExplanation.UnavailableText);

            try
            {
                Range start = worksheet.Range[startCell];
                string startFormula = start.Formula as string;
                if (string.IsNullOrWhiteSpace(startFormula))
                    return Miss($"{SheetCell(sheetName, startCell)}に数式がありません。");

                int startRow = start.Row;
                if (!isOk(startFormula, startRow))
                {
                    explainStart(sheetName, startCell, startFormula);
                    return false;
                }

                var missingOrWrong = new List<string>();
                foreach (Range cell in worksheet.Range[fillRangeAddress].Cells)
                {
                    string formula = cell.Formula as string;
                    if (!isOk(formula, cell.Row))
                        missingOrWrong.Add(cell.Address[false, false]);
                }

                if (missingOrWrong.Count == 0)
                    return true;

                return Miss($"シート「{sheetName}」の{JoinNames(missingOrWrong)}に正しい数式がありません。");
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // ===== formula matchers =====
        private static bool IsIfAriNashi(string formula, int row)
        {
            string n = NormalizeFormula(formula ?? "");
            if (n.Contains("$")) return false;
            return n.Contains("IF(")
                && n.Contains($"F{row}>5")
                && n.Contains("あり")
                && n.Contains("なし");
        }

        private static void ExplainIfAriNashi(string sheet, string cell, string formula)
        {
            string n = NormalizeFormula(formula ?? "");
            if (!n.Contains("IF("))
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}にIF関数がありません。");
            else
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の引数が「{Quote(DescribeFormulaArgs(formula, "IF"))}」になっています。");
            if (n.Contains("$"))
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の数式に「$」が含まれています。");
        }

        private static bool IsIfKyori(string formula, int row)
        {
            string n = NormalizeFormula(formula ?? "");
            if (n.Contains("$")) return false;
            return n.Contains("IF(")
                && n.Contains($"E{row}>=300")
                && n.Contains("10000")
                && n.Contains("5000");
        }

        private static void ExplainIfKyori(string sheet, string cell, string formula)
        {
            string n = NormalizeFormula(formula ?? "");
            if (!n.Contains("IF("))
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}にIF関数がありません。");
            else
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の引数が「{Quote(DescribeFormulaArgs(formula, "IF"))}」になっています。");
            if (n.Contains("$"))
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の数式に「$」が含まれています。");
        }

        private static bool IsIfStock(string formula, int row)
        {
            string n = NormalizeFormula(formula ?? "");
            if (n.Contains("$")) return false;
            // 解答手順は G4 で F5 を参照するため、Fn と F(n+1) の両方を許容
            bool cond = n.Contains($"F{row}<=13%") || n.Contains($"F{row}<=0.13")
                || n.Contains($"F{row + 1}<=13%") || n.Contains($"F{row + 1}<=0.13");
            return n.Contains("IF(")
                && cond
                && n.Contains("在庫を補充")
                && n.Contains("\"\"");
        }

        private static void ExplainIfStock(string sheet, string cell, string formula)
        {
            string n = NormalizeFormula(formula ?? "");
            if (!n.Contains("IF("))
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}にIF関数がありません。");
            else
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の引数が「{Quote(DescribeFormulaArgs(formula, "IF"))}」になっています。");
            if (n.Contains("$"))
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の数式に「$」が含まれています。");
        }

        private static bool IsTaxPrice(string formula, int row)
        {
            string n = NormalizeFormula(formula ?? "");
            return n.Contains($"E{row}")
                && n.Contains("$K$4")
                && n.Contains("*");
        }

        private static void ExplainTaxPrice(string sheet, string cell, string formula, int row)
        {
            string n = NormalizeFormula(formula ?? "");
            if (!n.Contains("*") && !n.Contains($"E{row}") && !n.Contains("$K$4"))
            {
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}の数式が「{Quote(formula)}」になっています。");
                return;
            }
            if (!n.Contains($"E{row}"))
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}が単価（E{row}）を参照していません。");
            if (!n.Contains("$K$4"))
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}が税率（$K$4）を参照していません。");
            if (!n.Contains("*"))
                ExcelScoreExplanation.Note($"{SheetCell(sheet, cell)}が掛け算の数式になっていません。");
        }

        // ===== COM helpers =====
        private bool TryOpenSheet(string filePath, string sheetName, out Workbook workbook, out Worksheet worksheet)
        {
            workbook = null;
            worksheet = null;
            Application excelApp;
            if (!TryGetWorkbook(filePath, out excelApp, out workbook))
                return false;
            worksheet = FindWorksheet(workbook, sheetName);
            return worksheet != null;
        }

        private bool TryGetWorkbook(string filePath, out Application excelApp, out Workbook workbook)
        {
            excelApp = null;
            workbook = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }

                string fileName = Path.GetFileName(filePath);
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private string GetCurrentExcelFilePath()
        {
            try
            {
                Application excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                return excelApp.ActiveWorkbook?.FullName;
            }
            catch
            {
                return null;
            }
        }

        private Worksheet FindWorksheet(Workbook workbook, string worksheetName)
        {
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase))
                    return ws;
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
