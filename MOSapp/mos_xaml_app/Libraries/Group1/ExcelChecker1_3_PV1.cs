using System;
using Excel = Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_3_PV1
    {
        public bool CheckTask_1_3_01() => Run(CheckTask_1_3_01_Impl);
        public bool CheckTask_1_3_02() => Run(CheckTask_1_3_02_Impl);
        public bool CheckTask_1_3_03() => Run(CheckTask_1_3_03_Impl);
        public bool CheckTask_1_3_04() => Run(CheckTask_1_3_04_Impl);
        public bool CheckTask_1_3_05() => Run(CheckTask_1_3_05_Impl);
        public bool CheckTask_1_3_06() => Run(CheckTask_1_3_06_Impl);

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

        private bool CheckTask_1_3_01_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "申込名簿");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            Excel.ListObject table = VariantCheckSupport.FirstTable(sheet);
            Excel.Range firstCell = table != null ? sheet.Range["K5"] : sheet.Range["K5"];
            string formula = VariantCheckSupport.CellFormula(firstCell);
            bool ok = VariantCheckSupport.NormalizeFormula(formula).Contains("市町村")
                && VariantCheckSupport.NormalizeFormula(formula).Contains("町名")
                && VariantCheckSupport.NormalizeFormula(formula).Contains("番地");
            if (!ok)
                ok = VariantCheckSupport.FormulaContainsAll(formula, "&");
            return ok ? true : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_3_02_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "申込名簿");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            sheet.Activate();
            Excel.Window window = app.ActiveWindow;
            if (!window.FreezePanes)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            return window.SplitRow == 4
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_3_03_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "申込名簿");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            Excel.ListObject table = VariantCheckSupport.FirstTable(sheet);
            Excel.Sort sort = table != null ? table.Sort : sheet.Sort;
            if (sort == null || sort.SortFields.Count < 3)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);

            return HasSort(sort, 1, "会員種別", Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlAscending)
                && HasSort(sort, 2, "テーマ", Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending)
                && HasSort(sort, 3, "よみ", Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlAscending)
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private static bool HasSort(
            Excel.Sort sort,
            int index,
            string header,
            Excel.XlSortOn sortOn,
            Excel.XlSortOrder order)
        {
            try
            {
                Excel.SortField field = sort.SortFields[index];
                string key = Convert.ToString(field.Key?.Cells[1, 1]?.Value2) ?? string.Empty;
                string address = VariantCheckSupport.NormalizeAddress(field.Key?.Address);
                bool nameOk = key.Contains(header) || address.Length > 0;
                return nameOk && field.Order == order;
            }
            catch
            {
                return false;
            }
        }

        private bool CheckTask_1_3_04_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "申込名簿");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            Excel.Range used = sheet.UsedRange;
            object[,] values = used.Value2 as object[,];
            if (values == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    string text = Convert.ToString(values[r, c]) ?? string.Empty;
                    if (text.Replace(" ", "　").Contains("中村　玲") || text.Replace("　", " ").Contains("中村 玲"))
                    {
                        for (int look = 1; look <= cols; look++)
                        {
                            string neighbor = Convert.ToString(values[r, look]) ?? string.Empty;
                            if (neighbor.Contains("経営企画部"))
                                return true;
                        }
                    }
                }
            }

            return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_3_05_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "申込名簿");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            foreach (Excel.Hyperlink link in sheet.Hyperlinks)
            {
                string address = link.Address ?? string.Empty;
                string range = VariantCheckSupport.NormalizeAddress(link.Range?.Address);
                if (range == "M4" && address.IndexOf("aoba-setsumeikai", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_3_06_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "会場一覧");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            Excel.ListObject table = VariantCheckSupport.FirstTable(sheet);
            if (table == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            return table.ShowHeaders
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }
    }
}
