using System;
using Excel = Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_4_PV1
    {
        private const int Style4ChartStyleId = 205;

        public bool CheckTask_1_4_01() => Run(CheckTask_1_4_01_Impl);
        public bool CheckTask_1_4_02() => Run(CheckTask_1_4_02_Impl);
        public bool CheckTask_1_4_03() => Run(CheckTask_1_4_03_Impl);
        public bool CheckTask_1_4_04() => Run(CheckTask_1_4_04_Impl);
        public bool CheckTask_1_4_05() => Run(CheckTask_1_4_05_Impl);
        public bool CheckTask_1_4_06() => Run(CheckTask_1_4_06_Impl);

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

        private bool CheckTask_1_4_01_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "月次利用");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            Excel.Range range = sheet.Range["B5:B20"];
            bool merged = false;
            try { merged = Convert.ToBoolean(range.MergeCells); } catch { }
            Excel.XlHAlign align = Excel.XlHAlign.xlHAlignGeneral;
            try { align = (Excel.XlHAlign)range.HorizontalAlignment; } catch { }
            bool centered = align == Excel.XlHAlign.xlHAlignCenter;
            return merged && centered
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_4_02_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "月次利用");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            double[] expected = { 7.88, 7.88, 5.38, 5.38 };
            for (int i = 0; i < expected.Length; i++)
            {
                double width = Convert.ToDouble(sheet.Columns[i + 5].ColumnWidth);
                if (Math.Abs(width - expected[i]) > 0.1)
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }

            return true;
        }

        private bool CheckTask_1_4_03_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Chart chart = FindStyleTargetChart(workbook, out string location);
            if (chart == null)
            {
                VariantCheckSupport.Diag("4-3 chart=null location=" + location);
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }

            object styleObj = chart.ChartStyle;
            if (styleObj == null || !int.TryParse(Convert.ToString(styleObj), out int styleId))
            {
                VariantCheckSupport.Diag("4-3 ChartStyle parse=fail location=" + location
                    + " value=" + VariantCheckSupport.TryRead(() => chart.ChartStyle));
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }

            bool styleOk = styleId == Style4ChartStyleId;
            VariantCheckSupport.Diag($"4-3 location={location} ChartStyle={styleId} expected={Style4ChartStyleId} styleOk={styleOk}");
            return styleOk
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private static Excel.Chart FindStyleTargetChart(Excel.Workbook workbook, out string location)
        {
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "月次利用");
            Excel.Chart embedded = sheet != null ? VariantCheckSupport.FindChart(sheet) : null;
            if (embedded != null)
            {
                location = "月次利用";
                return embedded;
            }

            if (workbook != null)
            {
                foreach (Excel.Chart chartSheet in workbook.Charts)
                {
                    if (string.Equals(chartSheet.Name, "拠点別利用グラフ", StringComparison.OrdinalIgnoreCase))
                    {
                        location = "拠点別利用グラフ";
                        return chartSheet;
                    }
                }
            }

            location = sheet == null ? "null" : "月次利用+no-chartsheet";
            return null;
        }

        private bool CheckTask_1_4_04_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            if (workbook == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            foreach (Excel.Chart chart in workbook.Charts)
            {
                if (string.Equals(chart.Name, "拠点別利用グラフ", StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_4_05_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "拠点別利用");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            string name = VariantCheckSupport.CellText(sheet.Range["H12"]);
            string region = VariantCheckSupport.CellText(sheet.Range["I12"]);
            string count = VariantCheckSupport.CellText(sheet.Range["J12"]);
            string hours = VariantCheckSupport.CellText(sheet.Range["K12"]);
            string amount = VariantCheckSupport.CellText(sheet.Range["L12"]);
            bool ok = name.Contains("千葉センター")
                && region.Contains("関東")
                && count.Replace(",", "").Contains("14")
                && hours.Replace(",", "").Contains("1770")
                && amount.Replace(",", "").Contains("1960000");
            VariantCheckSupport.Diag(
                $"4-5 H12:L12 name={name} region={region} count={count} hours={hours} amount={amount} ok={ok}");
            return ok
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_4_06_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            if (workbook == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            int sheetCount = workbook.Worksheets.Count;
            Excel.Worksheet last = workbook.Worksheets[sheetCount];
            string name = last.Name ?? string.Empty;
            if (name == "月次利用" || name == "拠点別利用")
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);

            bool hasTable = last.ListObjects.Count > 0;
            string a1 = VariantCheckSupport.CellText(last.Cells[1, 1]);
            return (hasTable || !string.IsNullOrWhiteSpace(a1))
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }
    }
}
