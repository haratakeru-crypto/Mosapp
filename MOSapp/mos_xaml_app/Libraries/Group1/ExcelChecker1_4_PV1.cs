using System;
using System.IO;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_4_PV1
    {
        private const int Style4ChartStyleId = 205;
        private const int Style4C14StyleId = 102;

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

            int comStyleId = -1;
            object styleObj = chart.ChartStyle;
            if (styleObj != null)
                int.TryParse(Convert.ToString(styleObj), out comStyleId);

            bool comOk = comStyleId == Style4ChartStyleId;
            string xmlDetail = "skip";
            bool xmlOk = false;
            if (!comOk)
                xmlOk = TryPersistedStyle4(workbook, filePath, out xmlDetail);

            VariantCheckSupport.Diag(
                $"4-3 location={location} ChartStyle={comStyleId} expected={Style4ChartStyleId} comOk={comOk} xml={xmlDetail} xmlOk={xmlOk}");
            return (comOk || xmlOk)
                ? true
                : VariantCheckSupport.Miss(ExcelScoreExplanation.RequirementMissText);
        }

        /// <summary>
        /// スタイル4適用直後の ChartStyle は 205 だが、保存して開き直すと COM は 2 に戻る。
        /// ディスク上の ChartStyle パーツ id="205" と c14:style val="102" で判定する。
        /// SaveCopyAs は開いているブックの COM 状態（2）を書いてしまうため使わない。
        /// </summary>
        private static bool TryPersistedStyle4(Excel.Workbook workbook, string filePath, out string detail)
        {
            string sourcePath = null;
            try { sourcePath = workbook?.FullName; } catch { }
            if (string.IsNullOrEmpty(sourcePath))
                sourcePath = filePath;

            if (TryReadPersistedStyle4FromPath(sourcePath, out detail))
                return true;

            if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath))
                return false;
            if (detail != null && !detail.StartsWith("err:", StringComparison.Ordinal)
                && !string.Equals(detail, "no-file", StringComparison.Ordinal))
            {
                return false;
            }

            string tempPath = Path.Combine(
                Path.GetTempPath(),
                "mos_pv1_4_3_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                File.Copy(sourcePath, tempPath, true);
                return TryReadPersistedStyle4FromPath(tempPath, out detail);
            }
            catch (Exception ex)
            {
                detail = "copy-err:" + ex.GetType().Name;
                return false;
            }
            finally
            {
                try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
            }
        }

        private static bool TryReadPersistedStyle4FromPath(string path, out string detail)
        {
            detail = "none";
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                detail = "no-file";
                return false;
            }

            try
            {
                using (FileStream stream = new FileStream(
                    path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (Package package = Package.Open(stream, FileMode.Open, FileAccess.Read))
                {
                    int partId = -1;
                    int c14 = -1;
                    foreach (PackagePart part in package.GetParts())
                    {
                        string uri = part.Uri != null ? part.Uri.OriginalString : string.Empty;
                        if (!uri.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                            continue;

                        bool isStylePart = uri.IndexOf("/xl/charts/style", StringComparison.OrdinalIgnoreCase) >= 0;
                        bool isChartPart = uri.IndexOf("/xl/charts/chart", StringComparison.OrdinalIgnoreCase) >= 0;
                        if (!isStylePart && !isChartPart)
                            continue;

                        string xml;
                        using (Stream partStream = part.GetStream(FileMode.Open, FileAccess.Read))
                        using (var reader = new StreamReader(partStream))
                            xml = reader.ReadToEnd();

                        if (isStylePart)
                        {
                            int foundId;
                            TryReadChartStylePartId(xml, out foundId);
                            if (foundId > 0)
                                partId = foundId;
                            if (foundId == Style4ChartStyleId)
                            {
                                detail = "partId=" + foundId;
                                return true;
                            }
                        }

                        if (isChartPart)
                        {
                            int foundC14;
                            TryReadC14StyleId(xml, out foundC14);
                            if (foundC14 > 0)
                                c14 = foundC14;
                            if (foundC14 == Style4C14StyleId)
                            {
                                detail = "c14=" + foundC14;
                                return true;
                            }
                        }
                    }

                    detail = "partId=" + partId + " c14=" + c14;
                    return false;
                }
            }
            catch (Exception ex)
            {
                detail = "err:" + ex.GetType().Name;
                return false;
            }
        }

        private static bool TryReadChartStylePartId(string xml, out int chartStyleId)
        {
            chartStyleId = -1;
            if (string.IsNullOrEmpty(xml))
                return false;

            Match partMatch = Regex.Match(
                xml, @"chartStyle\b[^>]*\bid\s*=\s*""(\d+)""", RegexOptions.IgnoreCase);
            if (!partMatch.Success)
                return false;
            return int.TryParse(partMatch.Groups[1].Value, out chartStyleId);
        }

        private static bool TryReadC14StyleId(string xml, out int c14StyleId)
        {
            c14StyleId = -1;
            if (string.IsNullOrEmpty(xml))
                return false;

            Match c14Match = Regex.Match(
                xml, @"c14:style\s+val\s*=\s*""(\d+)""", RegexOptions.IgnoreCase);
            if (!c14Match.Success)
                return false;
            return int.TryParse(c14Match.Groups[1].Value, out c14StyleId);
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
