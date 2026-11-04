using System;
using Excel = Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_2_PV1
    {
        public bool CheckTask_1_2_01() => Run(CheckTask_1_2_01_Impl);
        public bool CheckTask_1_2_02() => Run(CheckTask_1_2_02_Impl);
        public bool CheckTask_1_2_03() => Run(CheckTask_1_2_03_Impl);
        public bool CheckTask_1_2_04() => Run(CheckTask_1_2_04_Impl);
        public bool CheckTask_1_2_05() => Run(CheckTask_1_2_05_Impl);
        public bool CheckTask_1_2_06() => Run(CheckTask_1_2_06_Impl);

        private bool Run(Func<string, bool> impl)
        {
            try
            {
                Excel.Application app = VariantCheckSupport.GetApp();
                string path = app.ActiveWorkbook?.FullName;
                if (string.IsNullOrEmpty(path))
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
                return impl(path);
            }
            catch
            {
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        private bool CheckTask_1_2_01_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "収支計算書");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            try
            {
                return sheet.Cells.FormatConditions.Count == 0
                    ? true
                    : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }
            catch
            {
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        private bool CheckTask_1_2_02_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "収支分析");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            string formula = VariantCheckSupport.CellFormula(sheet.Range["F5"]);
            bool ok = VariantCheckSupport.FormulaContainsAll(formula, "SORT", "B5:C22")
                && (VariantCheckSupport.NormalizeFormula(formula).Contains("2,-1")
                    || VariantCheckSupport.NormalizeFormula(formula).Contains("2,-1)"));
            return ok ? true : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_2_03_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "収支計算書");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            for (int row = 6; row <= 9; row++)
            {
                string formula = VariantCheckSupport.CellFormula(sheet.Cells[row, 6]);
                if (!VariantCheckSupport.FormulaContainsAll(formula, "SUM", "C" + row, "E" + row))
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }

            return true;
        }

        private bool CheckTask_1_2_04_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "収支計算書");
            if (workbook == null || sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            string headerA = VariantCheckSupport.CellText(sheet.Range["A32"]);
            string headerB = VariantCheckSupport.CellText(sheet.Range["B32"]);
            if (string.IsNullOrEmpty(headerA) || string.IsNullOrEmpty(headerB))
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);

            bool nameA = false;
            bool nameB = false;
            foreach (Excel.Name name in workbook.Names)
            {
                try
                {
                    Excel.Range range = name.RefersToRange;
                    if (range == null || !string.Equals(range.Worksheet.Name, "収支計算書", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string address = VariantCheckSupport.NormalizeAddress(range.Address);
                    string shortName = name.Name;
                    int bang = shortName.LastIndexOf('!');
                    if (bang >= 0)
                        shortName = shortName.Substring(bang + 1);

                    if (address == "A33" && NameMatchesHeader(shortName, headerA))
                        nameA = true;
                    if (address == "B33" && NameMatchesHeader(shortName, headerB))
                        nameB = true;
                }
                catch
                {
                }
            }

            return nameA && nameB
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private static bool NameMatchesHeader(string definedName, string header)
        {
            if (string.IsNullOrEmpty(definedName) || string.IsNullOrEmpty(header))
                return false;
            if (string.Equals(definedName, header, StringComparison.OrdinalIgnoreCase))
                return true;

            string excelStyle = header.Replace(' ', '_').Replace('\u3000', '_');
            return string.Equals(definedName, excelStyle, StringComparison.OrdinalIgnoreCase);
        }

        private bool CheckTask_1_2_05_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "収支計算書");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            string footer = Convert.ToString(sheet.PageSetup.RightFooter) ?? string.Empty;
            bool hasSheetName = footer.IndexOf("&A", StringComparison.OrdinalIgnoreCase) >= 0
                || footer.IndexOf("収支計算書", StringComparison.Ordinal) >= 0;
            return hasSheetName ? true : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_2_06_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "収支分析");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            return sheet.ListObjects.Count == 0
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }
    }
}
