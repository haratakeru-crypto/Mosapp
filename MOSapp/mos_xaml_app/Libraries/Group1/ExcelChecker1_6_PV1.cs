using System;
using Excel = Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_6_PV1
    {
        public bool CheckTask_1_6_01() => Run(CheckTask_1_6_01_Impl);
        public bool CheckTask_1_6_02() => Run(CheckTask_1_6_02_Impl);
        public bool CheckTask_1_6_03() => Run(CheckTask_1_6_03_Impl);
        public bool CheckTask_1_6_04() => Run(CheckTask_1_6_04_Impl);
        public bool CheckTask_1_6_05() => Run(CheckTask_1_6_05_Impl);

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

        private bool CheckTask_1_6_01_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "価格対照");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            string a8 = VariantCheckSupport.CellText(sheet.Range["A8"]);
            string b8 = VariantCheckSupport.CellText(sheet.Range["B8"]);
            string originalB3 = VariantCheckSupport.CellText(sheet.Range["B3"]);
            bool transposed = !string.IsNullOrWhiteSpace(a8)
                && !string.Equals(a8, originalB3, StringComparison.Ordinal)
                || (!string.IsNullOrWhiteSpace(a8) && !string.IsNullOrWhiteSpace(b8));
            bool sourceKept = !string.IsNullOrWhiteSpace(VariantCheckSupport.CellText(sheet.Range["B3"]));
            return transposed && sourceKept
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_6_02_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "受注実績");
            Excel.Chart chart = sheet != null ? VariantCheckSupport.FindChart(sheet) : null;
            if (chart == null)
            {
                VariantCheckSupport.Diag("6-2 chart=null sheet=" + (sheet == null ? "null" : sheet.Name));
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }

            try
            {
                Excel.SeriesCollection seriesColl = (Excel.SeriesCollection)chart.SeriesCollection();
                int count = seriesColl.Count;
                VariantCheckSupport.Diag($"6-2 sheet={sheet.Name} seriesCount={count}");
                for (int i = 1; i <= count; i++)
                {
                    Excel.Series series = seriesColl.Item(i);
                    string name = VariantCheckSupport.TryRead(() => series.Name);
                    string formula = VariantCheckSupport.TryRead(() => series.Formula);
                    string values = DescribeSeriesRange(() => series.Values);
                    string xValues = DescribeSeriesRange(() => series.XValues);
                    string formulaNorm = VariantCheckSupport.NormalizeAddress(formula ?? string.Empty);
                    bool hasC4C21 = formulaNorm.IndexOf("C4:C21", StringComparison.OrdinalIgnoreCase) >= 0;
                    bool hasSalad = (formula ?? string.Empty).IndexOf("サラダ", StringComparison.Ordinal) >= 0;
                    VariantCheckSupport.Diag(
                        $"6-2 s{i} Name={name} Formula={formula} Values={values} XValues={xValues} formulaHasC4C21={hasC4C21} formulaHasSalad={hasSalad}");
                }

                bool rangeOk = false;
                for (int i = 1; i <= count; i++)
                {
                    Excel.Series series = seriesColl.Item(i);
                    string formula = Convert.ToString(series.Formula) ?? string.Empty;
                    string formulaNorm = VariantCheckSupport.NormalizeAddress(formula);
                    if (formulaNorm.IndexOf("C4:C21", StringComparison.OrdinalIgnoreCase) >= 0
                        || formula.IndexOf("サラダ", StringComparison.Ordinal) >= 0)
                    {
                        rangeOk = true;
                        break;
                    }
                }

                VariantCheckSupport.Diag($"6-2 check rangeOk={rangeOk}");
                return rangeOk
                    ? true
                    : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }
            catch (Exception ex)
            {
                VariantCheckSupport.Diag("6-2 exception " + ex.Message);
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        private static string DescribeSeriesRange(Func<object> getter)
        {
            try
            {
                object value = getter();
                if (value == null)
                    return "<null>";
                if (value is Excel.Range range)
                {
                    int lastRow = range.Row + range.Rows.Count - 1;
                    int lastCol = range.Column + range.Columns.Count - 1;
                    return $"{range.Address} row={range.Row}-{lastRow} col={range.Column}-{lastCol}";
                }

                return Convert.ToString(value) ?? "<empty>";
            }
            catch (Exception ex)
            {
                return "ERR:" + ex.GetType().Name + ":" + ex.Message;
            }
        }

        private bool CheckTask_1_6_03_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "受注実績");
            Excel.Chart chart = sheet != null ? VariantCheckSupport.FindChart(sheet) : null;
            if (chart == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);

            string alt = string.Empty;
            try { alt = Convert.ToString(((dynamic)chart).AlternativeText) ?? string.Empty; } catch { }
            if (string.IsNullOrEmpty(alt))
            {
                try
                {
                    Excel.ChartObject chartObject = sheet.ChartObjects(1);
                    alt = Convert.ToString(((dynamic)chartObject.ShapeRange).AlternativeText) ?? string.Empty;
                }
                catch
                {
                }
            }

            return alt.Contains("6月の品目別受注件数")
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_6_04_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "受注実績");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            string formula = VariantCheckSupport.CellFormula(sheet.Range["B24"]);
            bool formulaOk = VariantCheckSupport.FormulaContainsAll(formula, "SEQUENCE", "3", "6", "101");
            if (formulaOk)
                return true;

            object start = sheet.Range["B24"].Value2;
            object end = sheet.Range["G26"].Value2;
            bool valuesOk = start != null && end != null
                && Convert.ToDouble(start) == 101
                && Convert.ToDouble(end) == 118;
            return valuesOk ? true : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_6_05_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "受注実績");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            for (int row = 19; row <= 20; row++)
            {
                Excel.Range cell = sheet.Cells[row, 4];
                bool strike = false;
                bool bold = false;
                bool italic = false;
                try { strike = Convert.ToBoolean(cell.Font.Strikethrough); } catch { }
                try { bold = Convert.ToBoolean(cell.Font.Bold); } catch { }
                try { italic = Convert.ToBoolean(cell.Font.Italic); } catch { }
                if (!strike || !bold || !italic)
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }

            return true;
        }
    }
}
