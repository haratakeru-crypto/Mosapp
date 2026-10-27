using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_1_PV1
    {
        public bool CheckTask_1_1_01() => Run(CheckTask_1_1_01_Impl);
        public bool CheckTask_1_1_02() => Run(CheckTask_1_1_02_Impl);
        public bool CheckTask_1_1_03() => Run(CheckTask_1_1_03_Impl);
        public bool CheckTask_1_1_04() => Run(CheckTask_1_1_04_Impl);
        public bool CheckTask_1_1_05() => Run(CheckTask_1_1_05_Impl);
        public bool CheckTask_1_1_06() => Run(CheckTask_1_1_06_Impl);

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

        private bool CheckTask_1_1_01_Impl(string filePath)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            Excel.Range start = null;
            Excel.Range range = null;
            try
            {
                app = VariantCheckSupport.GetApp();
                workbook = VariantCheckSupport.GetWorkbook(app, filePath);
                sheet = VariantCheckSupport.FindSheet(workbook, "6月出荷台帳");
                if (sheet == null)
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

                start = sheet.Range["A5"];
                range = sheet.Range["A5:A28"];
                if (string.IsNullOrWhiteSpace(VariantCheckSupport.CellText(start)))
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);

                return VariantCheckSupport.IsConsecutiveSerials(range)
                    ? true
                    : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }
            catch
            {
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                VariantCheckSupport.Release(range, start, sheet);
            }
        }

        private bool CheckTask_1_1_02_Impl(string filePath)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            try
            {
                app = VariantCheckSupport.GetApp();
                workbook = VariantCheckSupport.GetWorkbook(app, filePath);
                sheet = VariantCheckSupport.FindSheet(workbook, "6月出荷台帳");
                if (sheet == null)
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

                for (int row = 5; row <= 28; row++)
                {
                    Excel.Range cell = sheet.Cells[row, 4];
                    string formula = VariantCheckSupport.CellFormula(cell);
                    bool formulaOk = VariantCheckSupport.FormulaContainsAll(formula, "RIGHT", "C" + row, "4");
                    if (!formulaOk)
                    {
                        string lot = VariantCheckSupport.CellText(sheet.Cells[row, 3]);
                        string symbol = VariantCheckSupport.CellText(cell);
                        if (string.IsNullOrEmpty(lot) || lot.Length < 4 ||
                            !string.Equals(symbol, lot.Substring(lot.Length - 4), StringComparison.Ordinal))
                        {
                            return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
                        }
                    }
                }

                return true;
            }
            catch
            {
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                VariantCheckSupport.Release(sheet);
            }
        }

        private bool CheckTask_1_1_03_Impl(string filePath)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            try
            {
                app = VariantCheckSupport.GetApp();
                workbook = VariantCheckSupport.GetWorkbook(app, filePath);
                sheet = VariantCheckSupport.FindSheet(workbook, "6月出荷台帳");
                DumpDataBarConditions(filePath, workbook, sheet);
                if (sheet == null)
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

                Excel.FormatConditions conditions = sheet.Range["G5:G28"].FormatConditions;
                for (int i = 1; i <= conditions.Count; i++)
                {
                    dynamic condition = conditions[i];
                    try
                    {
                        if ((int)condition.Type != 4)
                            continue;

                        string applies = VariantCheckSupport.NormalizeAddress(Convert.ToString(condition.AppliesTo?.Address));
                        if (applies != "G5:G28")
                            continue;

                        int barColor = ReadDataBarOleColor(condition);

                        bool solid = false;
                        try { solid = Convert.ToInt32(condition.BarFillType) == 0 || Convert.ToInt32(condition.FillType) == 0; } catch { solid = true; }

                        if (IsBlueDataBar(barColor) && solid)
                            return true;
                    }
                    catch
                    {
                    }
                }

                return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }
            catch (Exception ex)
            {
                VariantCheckSupport.Diag("1-3 unavailable " + ex.Message);
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                VariantCheckSupport.Release(sheet);
            }
        }

        private static void DumpDataBarConditions(string filePath, Excel.Workbook workbook, Excel.Worksheet sheet)
        {
            string wbName = workbook == null ? "null" : VariantCheckSupport.TryRead(() => workbook.Name);
            string sheetName = sheet == null ? "null" : VariantCheckSupport.TryRead(() => sheet.Name);
            VariantCheckSupport.Diag($"1-3 dump file={filePath} workbook={wbName} sheet={sheetName}");
            if (sheet == null)
                return;

            DumpFormatConditionCollection("G5:G28", () => sheet.Range["G5:G28"].FormatConditions);
            DumpFormatConditionCollection("UsedRange", () => sheet.UsedRange.FormatConditions);
        }

        private static void DumpFormatConditionCollection(string label, Func<Excel.FormatConditions> getter)
        {
            Excel.FormatConditions conditions;
            try
            {
                conditions = getter();
            }
            catch (Exception ex)
            {
                VariantCheckSupport.Diag($"1-3 {label} FormatConditions get ERR {ex.Message}");
                return;
            }

            int count = 0;
            try { count = conditions.Count; }
            catch (Exception ex)
            {
                VariantCheckSupport.Diag($"1-3 {label} Count ERR {ex.Message}");
                return;
            }

            VariantCheckSupport.Diag($"1-3 {label} Count={count}");
            for (int i = 1; i <= count; i++)
            {
                dynamic condition;
                try { condition = conditions[i]; }
                catch (Exception ex)
                {
                    VariantCheckSupport.Diag($"1-3 {label}[{i}] get ERR {ex.Message}");
                    continue;
                }

                try
                {
                    string typeText = VariantCheckSupport.TryRead(() => condition.Type);
                    string appliesRaw = VariantCheckSupport.TryRead(() => condition.AppliesTo?.Address);
                    string appliesA1 = VariantCheckSupport.TryRead(() => condition.AppliesTo?.Address[false, false]);
                    string applies = VariantCheckSupport.NormalizeAddress(
                        appliesRaw != null && !appliesRaw.StartsWith("ERR:") && appliesRaw != "<null>"
                            ? appliesRaw
                            : string.Empty);
                    string barColorText = VariantCheckSupport.TryRead(() => condition.BarColor);
                    string formatColor = VariantCheckSupport.TryRead(() => condition.BarColor.Color);
                    string colorType = VariantCheckSupport.TryRead(() => condition.BarColor.ColorType);
                    string barColorRgb = VariantCheckSupport.TryRead(() => condition.BarColor.RGB);
                    string barTheme = VariantCheckSupport.TryRead(() => condition.BarColor.ThemeColor);
                    string barTint = VariantCheckSupport.TryRead(() => condition.BarColor.TintAndShade);
                    string colorText = VariantCheckSupport.TryRead(() => condition.Color);
                    string colorRgb = VariantCheckSupport.TryRead(() => condition.Color.RGB);
                    string barFillText = VariantCheckSupport.TryRead(() => condition.BarFillType);
                    string fillText = VariantCheckSupport.TryRead(() => condition.FillType);
                    string direction = VariantCheckSupport.TryRead(() => condition.Direction);
                    string showValue = VariantCheckSupport.TryRead(() => condition.ShowValue);

                    int type = 0;
                    int.TryParse(typeText, out type);
                    int barColor = ReadDataBarOleColor(condition);

                    bool solid = false;
                    try { solid = Convert.ToInt32(condition.BarFillType) == 0 || Convert.ToInt32(condition.FillType) == 0; } catch { solid = true; }

                    bool typeOk = type == 4;
                    bool rangeOk = applies == "G5:G28";
                    bool blueOk = IsBlueDataBar(barColor);
                    VariantCheckSupport.Diag(
                        $"1-3 {label}[{i}] Type={typeText} AppliesRaw={appliesRaw} AppliesA1={appliesA1} Applies={applies} " +
                        $"BarColor={barColorText} FormatColor={formatColor} ColorType={colorType} BarRGB={barColorRgb} Theme={barTheme} Tint={barTint} Color={colorText} ColorRGB={colorRgb} " +
                        $"BarFill={barFillText} Fill={fillText} Direction={direction} ShowValue={showValue} " +
                        $"parsedColor={barColor} rgb=({barColor & 0xFF},{(barColor >> 8) & 0xFF},{(barColor >> 16) & 0xFF}) " +
                        $"typeOk={typeOk} rangeOk={rangeOk} blueOk={blueOk} solid={solid} wouldPass={typeOk && rangeOk && blueOk && solid}");
                }
                catch (Exception ex)
                {
                    VariantCheckSupport.Diag($"1-3 {label}[{i}] inspect ERR {ex.Message}");
                }
            }
        }

        private static int ReadDataBarOleColor(dynamic condition)
        {
            try
            {
                object value = condition.BarColor.Color;
                if (value != null)
                    return Convert.ToInt32(value);
            }
            catch
            {
            }

            try { return Convert.ToInt32(condition.BarColor); } catch { }
            try { return Convert.ToInt32(condition.Color); } catch { }
            return -1;
        }

        private static bool IsBlueDataBar(int color)
        {
            if (color <= 0)
                return false;
            int r = color & 0xFF;
            int g = (color >> 8) & 0xFF;
            int b = (color >> 16) & 0xFF;
            const int expectedR = 99;
            const int expectedG = 142;
            const int expectedB = 198;
            const int tolerance = 15;
            return Math.Abs(r - expectedR) <= tolerance
                && Math.Abs(g - expectedG) <= tolerance
                && Math.Abs(b - expectedB) <= tolerance;
        }

        private bool CheckTask_1_1_04_Impl(string filePath)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            try
            {
                app = VariantCheckSupport.GetApp();
                workbook = VariantCheckSupport.GetWorkbook(app, filePath);
                sheet = VariantCheckSupport.FindSheet(workbook, "6月出荷台帳");
                if (sheet == null)
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

                Excel.Range range = sheet.Range["I5:I28"];
                object wrap = range.WrapText;
                if (wrap is bool allWrapped && allWrapped)
                    return true;

                for (int row = 5; row <= 28; row++)
                {
                    if (!Convert.ToBoolean(sheet.Cells[row, 9].WrapText))
                        return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
                }

                return true;
            }
            catch
            {
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                VariantCheckSupport.Release(sheet);
            }
        }

        private bool CheckTask_1_1_05_Impl(string filePath)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            try
            {
                app = VariantCheckSupport.GetApp();
                workbook = VariantCheckSupport.GetWorkbook(app, filePath);
                sheet = VariantCheckSupport.FindSheet(workbook, "担当者集計");
                if (sheet == null)
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

                Excel.Chart chart = VariantCheckSupport.FindChart(sheet);
                if (chart == null)
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);

                int type = Convert.ToInt32(chart.ChartType);
                bool is3DClustered = type == (int)Excel.XlChartType.xl3DColumnClustered || type == 54;
                if (!is3DClustered)
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);

                Excel.ChartObject chartObject = sheet.ChartObjects(1);
                Excel.Range topLeft = (Excel.Range)chartObject.TopLeftCell;
                Excel.Range bottomRight = (Excel.Range)chartObject.BottomRightCell;
                bool placed =
                    topLeft.Row >= 26 && topLeft.Column >= 2 &&
                    bottomRight.Row <= 40 && bottomRight.Column <= 12;
                return placed ? true : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }
            catch
            {
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                VariantCheckSupport.Release(sheet);
            }
        }

        private bool CheckTask_1_1_06_Impl(string filePath)
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            try
            {
                app = VariantCheckSupport.GetApp();
                workbook = VariantCheckSupport.GetWorkbook(app, filePath);
                sheet = VariantCheckSupport.FindSheet(workbook, "6月出荷台帳");
                if (sheet == null)
                    return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);

                Excel.PageSetup setup = sheet.PageSetup;
                bool landscape = setup.Orientation == Excel.XlPageOrientation.xlLandscape;
                string area = VariantCheckSupport.NormalizeAddress(setup.PrintArea);
                bool areaOk = area == "A1:I30";
                return landscape && areaOk
                    ? true
                    : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
            }
            catch
            {
                return VariantCheckSupport.Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                VariantCheckSupport.Release(sheet);
            }
        }
    }
}
