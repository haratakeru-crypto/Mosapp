using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_5
    {
        // タスク5-2の正解スタイルID（正解は347）
        // バージョン差異を考慮して複数を許可していますが、実質的には347が対象です
        private readonly int[] VALID_STYLE_IDS = { 208, 8, 209, 9, 287, 288, 347 };


        // ラッパーメソッド
        public bool CheckTask_1_5_01() => RunCheck(CheckTask_1_5_01_Impl, "Task 5-1 (Layout)");
        public bool CheckTask_1_5_02() => RunCheck(CheckTask_1_5_02_Impl, "Task 5-2 (Style)");
        public bool CheckTask_1_5_03() => RunCheck(CheckTask_1_5_03_Impl, "Task 5-3 (Data Range)");
        public bool CheckTask_1_5_04() => RunCheck(CheckTask_1_5_04_Impl, "Task 5-4 (Legend)");
        public bool CheckTask_1_5_05() => RunCheck(CheckTask_1_5_05_Impl, "Task 5-5 (Label Pos)");
        public bool CheckTask_1_5_06() => RunCheck(CheckTask_1_5_06_Impl, "Task 5-6 (Axis Title)");


        // 共通エラーハンドリング
        private bool RunCheck(Func<string, bool> checkImpl, string taskName)
        {
            try
            {
                Console.WriteLine($"[DEBUG] {taskName} called");
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return checkImpl(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Exception in {taskName}: {ex.Message}");
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }


        // ==========================================
        // タスク5-1: クイックレイアウト・グラフタイトル
        // ChartLayout は COM/VSTO から読めないため、円グラフのレイアウト1相当を
        // 「タイトル + 凡例なし + パーセント付きデータラベル」で代理判定する。
        // （レイアウト2/6は凡例あり、レイアウト5は％なしで区別）
        // ==========================================
        private bool CheckTask_1_5_01_Impl(string filePath)
        {
            return ProcessChartTask(filePath, "売上実績", (chart) =>
            {
                bool titleOk = false;
                if (!chart.HasTitle)
                {
                    ExcelScoreExplanation.Note("グラフタイトルがありません。");
                }
                else
                {
                    string title = chart.ChartTitle.Text.Trim();
                    Console.WriteLine($"[DEBUG] Current Title: '{title}'");
                    if (title.StartsWith("売上構成比"))
                        titleOk = true;
                    else
                        ExcelScoreExplanation.Note($"グラフタイトルが「{Quote(title)}」になっています。");
                }

                bool noLegend = !chart.HasLegend;
                bool hasPercentLabels = ChartHasPercentageDataLabels(chart);
                bool layoutOk = noLegend && hasPercentLabels;
                Console.WriteLine($"[DEBUG] Layout1 proxy (NoLegend={noLegend}, ShowPercentage={hasPercentLabels})");
                if (!layoutOk)
                    ExcelScoreExplanation.Note("グラフのレイアウトがレイアウト1になっていません。");

                if (titleOk && layoutOk)
                {
                    Console.WriteLine("[DEBUG] Task 5-1 Passed: Title and Layout1 proxy.");
                    return true;
                }

                return false;
            });
        }


        // ==========================================
        // タスク5-2: グラフスタイル・配色の変更
        // ==========================================
        private bool CheckTask_1_5_02_Impl(string filePath)
        {
            return ProcessChartTask(filePath, "商品別売上", (chart) =>
            {
                object styleObj = chart.ChartStyle;
                int styleId = -1;

                if (styleObj != null && int.TryParse(styleObj.ToString(), out styleId))
                {
                    Console.WriteLine($"[DEBUG] Current Chart Style ID: {styleId}");
                }

                bool isStyleCorrect = false;
                foreach (int validId in VALID_STYLE_IDS)
                {
                    if (styleId == validId) isStyleCorrect = true;
                }

                object colorObj = chart.ChartColor;
                int colorId = -1;
                if (colorObj != null && int.TryParse(colorObj.ToString(), out colorId))
                {
                    Console.WriteLine($"[DEBUG] Current Chart Color: {colorId}");
                }

                bool isColorCorrect = (colorId == 12);

                if (isStyleCorrect && isColorCorrect)
                {
                    Console.WriteLine("[DEBUG] Task 5-2 Passed: Style ID and Color matched.");
                    return true;
                }

                // ChartStyle / ChartColor の数値は画面の「スタイルN」「パレットN」と一致しないため、内部IDは出さない。
                if (!isStyleCorrect)
                    ExcelScoreExplanation.Note("グラフスタイルがスタイル12になっていません。");
                if (!isColorCorrect)
                    ExcelScoreExplanation.Note("グラフの配色がカラフルなパレット3になっていません。");
                return false;
            });
        }


        // ==========================================
        // タスク5-3: データの追加（範囲変更）・ラベル非表示
        // ==========================================
        private bool CheckTask_1_5_03_Impl(string filePath)
        {
            return ProcessChartTask(filePath, "商品別売上", (chart) =>
            {
                bool titleOk = true;
                if (chart.HasTitle)
                {
                    string title = chart.ChartTitle.Text.Trim();
                    if (title != "総計" && !string.IsNullOrEmpty(title))
                    {
                        titleOk = false;
                        ExcelScoreExplanation.Note($"グラフタイトルが「{Quote(title)}」に変わっています。");
                    }
                }

                bool layoutOk = !(chart.HasLegend && chart.Legend.Position == Excel.XlLegendPosition.xlLegendPositionRight);
                if (!layoutOk)
                    ExcelScoreExplanation.Note("グラフのレイアウトが変わっています（凡例が右にあります）。");

                Excel.SeriesCollection seriesColl = (Excel.SeriesCollection)chart.SeriesCollection();
                if (seriesColl.Count == 0)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                bool labelsHidden = true;
                foreach (Excel.Series s in seriesColl)
                {
                    if (s.HasDataLabels)
                    {
                        labelsHidden = false;
                        break;
                    }
                }
                if (!labelsHidden)
                    ExcelScoreExplanation.Note("データラベルが表示されたままです。");

                // 初期グラフにも行10は含まれることが多いため、$10 判定は使わない。
                // 解答手順どおり B 列まで広げて 1〜5 月が入っているかを見る。
                bool rangeExpanded = ChartSeriesIncludesColumnB(seriesColl);
                Console.WriteLine($"[DEBUG] Task 5-3 range includes column B: {rangeExpanded}");
                if (!rangeExpanded)
                    ExcelScoreExplanation.Note("グラフに1月から5月のデータが追加されていません。");

                if (titleOk && layoutOk && labelsHidden && rangeExpanded)
                {
                    Console.WriteLine("[DEBUG] Task 5-3 Passed: Range includes column B and labels hidden.");
                    return true;
                }

                return false;
            });
        }


        // ==========================================
        // タスク5-4: 凡例を上に表示
        // ==========================================
        private bool CheckTask_1_5_04_Impl(string filePath)
        {
            return ProcessChartTask(filePath, "商品別売上", (chart) =>
            {
                bool titleOk = true;
                if (chart.HasTitle)
                {
                    string title = chart.ChartTitle.Text.Trim();
                    if (title != "総計" && !string.IsNullOrEmpty(title))
                    {
                        titleOk = false;
                        ExcelScoreExplanation.Note($"グラフタイトルが「{Quote(title)}」に変わっています。");
                    }
                }

                bool legendOk = false;
                if (!chart.HasLegend)
                {
                    ExcelScoreExplanation.Note("凡例が表示されていません。");
                }
                else if (chart.Legend.Position == Excel.XlLegendPosition.xlLegendPositionTop)
                {
                    legendOk = true;
                }
                else
                {
                    ExcelScoreExplanation.Note($"凡例の位置が「{DescribeLegendPosition(chart.Legend.Position)}」になっています。");
                }

                if (titleOk && legendOk)
                {
                    Console.WriteLine("[DEBUG] Task 5-4 Passed: Legend is at Top.");
                    return true;
                }

                return false;
            });
        }


        // ==========================================
        // タスク5-5: データラベル「内部外側」
        // ==========================================
        private bool CheckTask_1_5_05_Impl(string filePath)
        {
            return ProcessChartTask(filePath, "月別売上", (chart) =>
            {
                Excel.SeriesCollection seriesColl = (Excel.SeriesCollection)chart.SeriesCollection();
                if (seriesColl.Count == 0)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Excel.Series series = seriesColl.Item(1);

                if (!series.HasDataLabels)
                    return Miss("データラベルが表示されていません。");

                Excel.DataLabels dataLabels = (Excel.DataLabels)series.DataLabels();

                if (dataLabels.Position == Excel.XlDataLabelPosition.xlLabelPositionInsideEnd)
                {
                    Console.WriteLine("[DEBUG] Task 5-5 Passed: Label position is InsideEnd.");
                    return true;
                }

                return Miss($"データラベルの位置が「{DescribeDataLabelPosition(dataLabels.Position)}」になっています。");
            });
        }


        // ==========================================
        // タスク5-6: 第1横軸ラベル（軸タイトル）
        // 月別売上は横棒グラフのため、画面の第1横軸は数値軸(xlValue)。
        // （項目軸 xlCategory は縦方向になる）
        // ==========================================
        private bool CheckTask_1_5_06_Impl(string filePath)
        {
            return ProcessChartTask(filePath, "月別売上", (chart) =>
            {
                string horizontalTitle = TryGetAxisTitle(chart, Excel.XlAxisType.xlValue);
                string verticalTitle = TryGetAxisTitle(chart, Excel.XlAxisType.xlCategory);

                if (!string.IsNullOrEmpty(horizontalTitle))
                {
                    if (horizontalTitle == "単位：円" || horizontalTitle.Contains("単位：円"))
                    {
                        Console.WriteLine($"[DEBUG] Task 5-6 Passed: Horizontal axis title '{horizontalTitle}'.");
                        return true;
                    }
                    return Miss($"第1横軸ラベルが「{Quote(horizontalTitle)}」になっています。");
                }

                // 横軸未設定。縦軸に何かある場合はそちらを理由に出す。
                if (!string.IsNullOrEmpty(verticalTitle))
                {
                    if (verticalTitle == "単位：円" || verticalTitle.Contains("単位：円"))
                        return Miss("第1横軸ではなく第1縦軸に設定されています。");
                    return Miss($"第1縦軸ラベルが設定されていて、「{Quote(verticalTitle)}」となっています。");
                }

                return Miss("第1横軸ラベル（単位：円）が表示されていません。");
            });
        }


        // ==========================================
        // 共通ヘルパー
        // ==========================================
        private bool ProcessChartTask(string filePath, string sheetName, Func<Excel.Chart, bool> checkLogic)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                try { excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Excel.Application { Visible = false }; }

                string fileName = Path.GetFileName(filePath);
                foreach (Excel.Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = ws;
                        break;
                    }
                }
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
                if (chartObjects.Count == 0)
                    return Miss($"シート「{sheetName}」にグラフがありません。");

                Excel.ChartObject chartObj = chartObjects.Item(1);
                return checkLogic(chartObj.Chart);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error in ProcessChartTask: {ex.Message}");
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private string GetCurrentExcelFilePath()
        {
            try
            {
                Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                return excelApp.ActiveWorkbook?.FullName;
            }
            catch { return null; }
        }

        private static bool Miss(string reason)
        {
            ExcelScoreExplanation.Note(reason);
            return false;
        }

        private static string TryGetAxisTitle(Excel.Chart chart, Excel.XlAxisType axisType)
        {
            try
            {
                Excel.Axis axis = (Excel.Axis)chart.Axes(axisType);
                if (axis != null && axis.HasTitle)
                {
                    string t = axis.AxisTitle.Text?.Trim();
                    if (!string.IsNullOrEmpty(t))
                        return t;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] TryGetAxisTitle({axisType}): {ex.Message}");
            }
            return null;
        }

        /// <summary>
        /// 系列のデータ範囲に B 列（1〜5月側）が含まれるか。
        /// </summary>
        private static bool ChartSeriesIncludesColumnB(Excel.SeriesCollection seriesColl)
        {
            if (seriesColl == null || seriesColl.Count == 0)
                return false;

            foreach (Excel.Series series in seriesColl)
            {
                try
                {
                    if (series.Values is Excel.Range valuesRange)
                    {
                        int startCol = valuesRange.Column;
                        int endCol = startCol + valuesRange.Columns.Count - 1;
                        if (startCol <= 2 && endCol >= 2)
                            return true;
                    }
                }
                catch
                {
                    // Values が Range でない場合は Formula へ
                }

                try
                {
                    string formula = series.Formula ?? "";
                    Console.WriteLine($"[DEBUG] Series Formula: {formula}");
                    if (formula.IndexOf("$B$", StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
                catch
                {
                    // 次の系列へ
                }
            }

            return false;
        }

        /// <summary>
        /// 系列にパーセント付きデータラベルがあるか（クイックレイアウト1の代理指標）。
        /// </summary>
        private static bool ChartHasPercentageDataLabels(Excel.Chart chart)
        {
            try
            {
                Excel.SeriesCollection seriesColl = (Excel.SeriesCollection)chart.SeriesCollection();
                if (seriesColl == null || seriesColl.Count == 0)
                    return false;

                foreach (Excel.Series series in seriesColl)
                {
                    try
                    {
                        if (!series.HasDataLabels)
                            continue;

                        Excel.DataLabels dataLabels = (Excel.DataLabels)series.DataLabels();
                        if (dataLabels != null && dataLabels.ShowPercentage)
                            return true;
                    }
                    catch
                    {
                        // 系列ごとに読めない場合は次へ
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] ChartHasPercentageDataLabels: {ex.Message}");
            }

            return false;
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

        private static string DescribeLegendPosition(Excel.XlLegendPosition position)
        {
            switch (position)
            {
                case Excel.XlLegendPosition.xlLegendPositionTop: return "上";
                case Excel.XlLegendPosition.xlLegendPositionBottom: return "下";
                case Excel.XlLegendPosition.xlLegendPositionLeft: return "左";
                case Excel.XlLegendPosition.xlLegendPositionRight: return "右";
                case Excel.XlLegendPosition.xlLegendPositionCorner: return "隅";
                default: return position.ToString();
            }
        }

        private static string DescribeDataLabelPosition(Excel.XlDataLabelPosition position)
        {
            switch (position)
            {
                case Excel.XlDataLabelPosition.xlLabelPositionInsideEnd: return "内側上端（内部外側）";
                case Excel.XlDataLabelPosition.xlLabelPositionInsideBase: return "内側下端";
                case Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd: return "外側上端";
                case Excel.XlDataLabelPosition.xlLabelPositionCenter: return "中央";
                case Excel.XlDataLabelPosition.xlLabelPositionBestFit: return "自動";
                case Excel.XlDataLabelPosition.xlLabelPositionLeft: return "左";
                case Excel.XlDataLabelPosition.xlLabelPositionRight: return "右";
                case Excel.XlDataLabelPosition.xlLabelPositionAbove: return "上";
                case Excel.XlDataLabelPosition.xlLabelPositionBelow: return "下";
                default: return position.ToString();
            }
        }
    }
}
