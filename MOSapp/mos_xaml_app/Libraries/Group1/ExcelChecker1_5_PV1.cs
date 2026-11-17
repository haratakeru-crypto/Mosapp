using System;
using Excel = Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_5_PV1
    {
        public bool CheckTask_1_5_01() => Run(CheckTask_1_5_01_Impl);
        public bool CheckTask_1_5_02() => Run(CheckTask_1_5_02_Impl);
        public bool CheckTask_1_5_03() => Run(CheckTask_1_5_03_Impl);
        public bool CheckTask_1_5_04() => Run(CheckTask_1_5_04_Impl);
        public bool CheckTask_1_5_05() => Run(CheckTask_1_5_05_Impl);
        public bool CheckTask_1_5_06() => Run(CheckTask_1_5_06_Impl);

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

        private bool CheckTask_1_5_01_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "出席名簿");
            Excel.ListObject table = sheet != null ? VariantCheckSupport.FirstTable(sheet) : null;
            if (table == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);

            return table.ShowTableStyleRowStripes
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_5_02_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "出席名簿");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            string formula = VariantCheckSupport.CellFormula(sheet.Range["B3"]);
            // 出欠列の空白個数。正答例: =COUNTBLANK(出席一覧[出欠])
            bool ok = VariantCheckSupport.FormulaContainsAll(formula, "COUNTBLANK", "出欠");
            return ok ? true : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_5_03_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "出席名簿");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            string formula = VariantCheckSupport.CellFormula(sheet.Range["F6"]);
            string normalized = VariantCheckSupport.NormalizeFormula(formula);
            // NormalizeFormula は引用符を落とすため、空欄判定 "" / ISBLANK は元数式側で見る。
            // 正答例: =IF([@出欠]="","送付しない","当日配布")
            string rawCompact = (formula ?? string.Empty).Replace(" ", "");
            bool hasBlankCond = rawCompact.IndexOf("\"\"", StringComparison.Ordinal) >= 0
                || normalized.Contains("ISBLANK");
            bool ok = normalized.Contains("IF")
                && normalized.Contains("出欠")
                && normalized.Contains("送付しない")
                && normalized.Contains("当日配布")
                && hasBlankCond;
            return ok ? true : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_5_04_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "出席名簿");
            Excel.ListObject table = sheet != null ? VariantCheckSupport.FirstTable(sheet) : null;
            if (table == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);

            return table.ShowTableStyleLastColumn
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_5_05_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            Excel.Worksheet sheet = VariantCheckSupport.FindSheet(workbook, "日程表");
            if (sheet == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            int align = 0;
            try
            {
                align = Convert.ToInt32(sheet.Range["A1"].HorizontalAlignment);
            }
            catch
            {
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
            }

            int rightAlign = 0;
            try
            {
                rightAlign = Convert.ToInt32(sheet.Range["F1"].HorizontalAlignment);
            }
            catch { }

            bool ok = align == (int)Excel.XlHAlign.xlHAlignCenterAcrossSelection
                && rightAlign != (int)Excel.XlHAlign.xlHAlignCenterAcrossSelection;
            return ok
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        private bool CheckTask_1_5_06_Impl(string filePath)
        {
            Excel.Application app = VariantCheckSupport.GetApp();
            Excel.Workbook workbook = VariantCheckSupport.GetWorkbook(app, filePath);
            if (workbook == null)
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

            string category = VariantCheckSupport.TryGetBuiltinProperty(workbook, "Category");
            bool ok = category.Contains("定時説明会")
                && category.Contains("株主")
                && category.Contains("2026年6月");
            return ok ? true : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }
    }
}
