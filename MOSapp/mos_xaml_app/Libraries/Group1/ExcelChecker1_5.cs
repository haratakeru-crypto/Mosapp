using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

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
                if (string.IsNullOrEmpty(filePath)) return false;
                return checkImpl(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Exception in {taskName}: {ex.Message}");
                return false;
            }
        }


        // ==========================================
        // タスク5-1: クイックレイアウト・グラフタイトル
        // ==========================================
        private bool CheckTask_1_5_01_Impl(string filePath)
        {
            return ProcessChartTask(filePath, "売上実績", (chart) =>
            {
                if (!chart.HasTitle)
                {
                    Console.WriteLine("[DEBUG] Failed: Chart has no title.");
                    return false;
                }

                string title = chart.ChartTitle.Text.Trim();
                Console.WriteLine($"[DEBUG] Current Title: '{title}'");
                
                if (!title.StartsWith("売上構成比"))
                {
                    Console.WriteLine($"[DEBUG] Failed: Title mismatch. Expected to start with: '売上構成比'");
                    return false;
                }

                bool hasLegend = chart.HasLegend;
                Console.WriteLine($"[DEBUG] HasLegend: {hasLegend}");
                
                if (hasLegend)
                {
                    Console.WriteLine($"[DEBUG] Legend Position: {chart.Legend.Position}");
                    if (chart.Legend.Position == Excel.XlLegendPosition.xlLegendPositionRight)
                    {
                        Console.WriteLine("[DEBUG] Task 5-1 Passed: Title and Legend position correct.");
                        return true;
                    }
                }

                Console.WriteLine("[DEBUG] Task 5-1 Passed: Title is correct.");
                return true;
            });
        }


        // ==========================================
        // タスク5-2: グラフスタイル・配色の変更
        // ==========================================
        private bool CheckTask_1_5_02_Impl(string filePath)
        {
            return ProcessChartTask(filePath, "商品別売上", (chart) =>
            {
                // 1. スタイルのチェック
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

                // 2. 配色 (ChartColor) のチェックを追加
                // 「カラフルなパレット3」は通常 ID 12
                object colorObj = chart.ChartColor;
                int colorId = -1;
                if (colorObj != null && int.TryParse(colorObj.ToString(), out colorId))
                {
                    Console.WriteLine($"[DEBUG] Current Chart Color: {colorId}");
                }
                
                // 厳密に 12 であることを要求
                bool isColorCorrect = (colorId == 12);

                if (isStyleCorrect && isColorCorrect)
                {
                    Console.WriteLine("[DEBUG] Task 5-2 Passed: Style ID and Color matched.");
                    return true;
                }

                if (!isStyleCorrect) Console.WriteLine("[DEBUG] Failed: Style ID mismatch.");
                if (!isColorCorrect) Console.WriteLine($"[DEBUG] Failed: Color mismatch. Expected 12, got {colorId}.");
                
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
                // レイアウト変更の検知
                // 1. タイトルが初期状態から変わっている場合
                if (chart.HasTitle)
                {
                    string title = chart.ChartTitle.Text.Trim();
                    if (title != "総計" && !string.IsNullOrEmpty(title))
                    {
                        Console.WriteLine($"[DEBUG] Failed: Chart title changed ('{title}'), likely due to layout change.");
                        return false;
                    }
                }

                // 2. 凡例が右側にある場合（レイアウト1）
                if (chart.HasLegend && chart.Legend.Position == Excel.XlLegendPosition.xlLegendPositionRight)
                {
                    Console.WriteLine("[DEBUG] Failed: Layout appears to be changed (Legend on right = Layout 1).");
                    return false;
                }

                // 1. データラベル非表示のチェック
                Excel.SeriesCollection seriesColl = (Excel.SeriesCollection)chart.SeriesCollection();
                if (seriesColl.Count == 0) return false;

                foreach (Excel.Series s in seriesColl)
                {
                    if (s.HasDataLabels)
                    {
                        Console.WriteLine("[DEBUG] Failed: Data Labels are visible.");
                        return false; 
                    }
                }

                // 2. データの範囲変更（追加）のチェック
                Excel.Series series1 = seriesColl.Item(1);
                string formula = series1.Formula;
                Console.WriteLine($"[DEBUG] Series Formula: {formula}");

                bool rangeExpanded = formula.Contains("$10") || formula.Contains(":10");

                if (rangeExpanded)
                {
                    Console.WriteLine("[DEBUG] Task 5-3 Passed: Range expanded ($10) and labels hidden.");
                    return true;
                }

                Console.WriteLine("[DEBUG] Failed: Range does not match expected expansion (Expected $10).");
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
                // レイアウト変更の検知（タイトルが初期状態から変わっている場合）
                if (chart.HasTitle)
                {
                    string title = chart.ChartTitle.Text.Trim();
                    // 初期タイトルは「総計」なので、それ以外に変わっている場合はレイアウト適用と判断
                    if (title != "総計" && !string.IsNullOrEmpty(title))
                    {
                        Console.WriteLine($"[DEBUG] Failed: Chart title changed ('{title}'), likely due to layout change.");
                        return false;
                    }
                }

                if (!chart.HasLegend)
                {
                    Console.WriteLine("[DEBUG] Failed: No legend found.");
                    return false;
                }

                if (chart.Legend.Position == Excel.XlLegendPosition.xlLegendPositionTop)
                {
                    Console.WriteLine("[DEBUG] Task 5-4 Passed: Legend is at Top.");
                    return true;
                }

                Console.WriteLine($"[DEBUG] Failed: Legend position is {chart.Legend.Position}.");
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
                if (seriesColl.Count == 0) return false;

                Excel.Series series = seriesColl.Item(1);

                if (!series.HasDataLabels)
                {
                    Console.WriteLine("[DEBUG] Failed: Data Labels not enabled.");
                    return false;
                }

                Excel.DataLabels dataLabels = (Excel.DataLabels)series.DataLabels();
                
                if (dataLabels.Position == Excel.XlDataLabelPosition.xlLabelPositionInsideEnd)
                {
                    Console.WriteLine("[DEBUG] Task 5-5 Passed: Label position is InsideEnd.");
                    return true;
                }

                Console.WriteLine($"[DEBUG] Failed: Label position is {dataLabels.Position}.");
                return false;
            });
        }


        // ==========================================
        // タスク5-6: 第1横軸ラベル（軸タイトル）
        // ==========================================
        private bool CheckTask_1_5_06_Impl(string filePath)
        {
            return ProcessChartTask(filePath, "月別売上", (chart) =>
            {
                // 軸のタイプを柔軟にチェック (xlCategory or xlValue)
                bool titleFound = false;
                string foundTitle = "";
                var axisTypes = new[] { Excel.XlAxisType.xlCategory, Excel.XlAxisType.xlValue };

                foreach (var type in axisTypes)
                {
                    try
                    {
                        Excel.Axis axis = (Excel.Axis)chart.Axes(type);
                        if (axis.HasTitle)
                        {
                            string t = axis.AxisTitle.Text.Trim();
                            // "単位：円" を含むか、完全一致か
                            if (t == "単位：円" || t.Contains("単位：円"))
                            {
                                titleFound = true;
                                foundTitle = t;
                                break;
                            }
                        }
                    }
                    catch
                    {
                        // 軸が存在しない場合等は無視
                    }
                }

                if (titleFound)
                {
                    Console.WriteLine($"[DEBUG] Task 5-6 Passed: Axis title '{foundTitle}' found.");
                    return true;
                }

                Console.WriteLine("[DEBUG] Failed: X-Axis (or any axis) has no matching title.");
                return false;
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
                if (workbook == null) return false;

                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = ws;
                        break;
                    }
                }
                if (worksheet == null) return false;

                Excel.ChartObjects chartObjects = (Excel.ChartObjects)worksheet.ChartObjects();
                if (chartObjects.Count == 0) return false;

                Excel.ChartObject chartObj = chartObjects.Item(1);
                return checkLogic(chartObj.Chart);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error in ProcessChartTask: {ex.Message}");
                return false;
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
    }
}
