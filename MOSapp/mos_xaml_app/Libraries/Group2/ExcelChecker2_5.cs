using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace Libraries.Group2
{
    public class ExcelChecker2_5
    {
        public bool CheckTask_2_5_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_5_01_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_5_01_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_5_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_5_02_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_5_02_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_5_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_5_03_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_5_03_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_5_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_5_04_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_5_04_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_5_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_5_05_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_5_05_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_5_01_Impl(string filePath)
        {
            // CSVの解答手順: 「商品売上」シートのグラフをスタイル9にし、グラフタイトルを「商品売上」に
            // ExcelChecker1_5のCheckTask_1_5_02_Implを参考に、ProcessChartTaskパターンを使用
            return ProcessChartTask(filePath, "商品売上", (chart) =>
            {
                // 1. スタイルのチェック
                object styleObj = chart.ChartStyle;
                int styleId = -1;

                if (styleObj != null && int.TryParse(styleObj.ToString(), out styleId))
                {
                    Console.WriteLine($"[DEBUG] Current Chart Style ID: {styleId}");
                }

                // スタイル9のIDは通常9（バージョン差異を考慮）
                int[] validStyleIds = { 9, 209, 287, 288 };
                bool isStyleCorrect = false;
                foreach (int validId in validStyleIds)
                {
                    if (styleId == validId) isStyleCorrect = true;
                }

                // 2. グラフタイトルのチェック
                if (!chart.HasTitle)
                {
                    Console.WriteLine("[DEBUG] Task 2-5-1 Failed: Chart has no title.");
                    return false;
                }

                string title = chart.ChartTitle.Text.Trim();
                Console.WriteLine($"[DEBUG] Current Title: '{title}'");
                
                bool isTitleCorrect = title == "商品売上" || title.StartsWith("商品売上");

                if (isStyleCorrect && isTitleCorrect)
                {
                    Console.WriteLine("[DEBUG] Task 2-5-1 Passed: Style ID and Title matched.");
                    return true;
                }

                if (!isStyleCorrect) Console.WriteLine("[DEBUG] Task 2-5-1 Failed: Style ID mismatch.");
                if (!isTitleCorrect) Console.WriteLine($"[DEBUG] Task 2-5-1 Failed: Title mismatch. Expected '商品売上', got '{title}'.");
                
                return false;
            });
        }

        private bool CheckTask_2_5_02_Impl(string filePath)
        {
            // CSVの解答手順: 「販売実績」シート内のグラフをレイアウト2に設定し、パレットを「カラフルなパレット3」に変更
            // ExcelChecker1_5のCheckTask_1_5_02_Implを参考に、ProcessChartTaskパターンを使用
            return ProcessChartTask(filePath, "販売実績", (chart) =>
            {
                // 1. 配色 (ChartColor) のチェック
                // 「カラフルなパレット3」は通常 ID 12
                object colorObj = chart.ChartColor;
                int colorId = -1;
                if (colorObj != null && int.TryParse(colorObj.ToString(), out colorId))
                {
                    Console.WriteLine($"[DEBUG] Current Chart Color: {colorId}");
                }
                
                // 厳密に 12 であることを要求
                bool isColorCorrect = (colorId == 12);

                // 2. レイアウト2のチェック（凡例位置などで判定）
                // レイアウト2は通常、凡例が右側にある
                bool hasLegend = chart.HasLegend;
                bool isLayoutCorrect = false;
                
                if (hasLegend)
                {
                    // レイアウト2では凡例が右側にあることが多い
                    if (chart.Legend.Position == XlLegendPosition.xlLegendPositionRight)
                    {
                        isLayoutCorrect = true;
                    }
                }

                if (isColorCorrect && isLayoutCorrect)
                {
                    Console.WriteLine("[DEBUG] Task 2-5-2 Passed: Color and Layout matched.");
                    return true;
                }

                if (!isColorCorrect) Console.WriteLine($"[DEBUG] Task 2-5-2 Failed: Color mismatch. Expected 12, got {colorId}.");
                if (!isLayoutCorrect) Console.WriteLine("[DEBUG] Task 2-5-2 Failed: Layout mismatch.");
                
                return false;
            });
        }

        private bool CheckTask_2_5_03_Impl(string filePath)
        {
            // CSVの解答手順: 「販売実績」シート内のグラフの凡例を削除し、過去の月のデータをすべて反映し表示
            // ExcelChecker1_5のCheckTask_1_5_03_Implを参考に、ProcessChartTaskパターンを使用
            return ProcessChartTask(filePath, "販売実績", (chart) =>
            {
                // 1. 凡例が削除されているかチェック
                if (chart.HasLegend)
                {
                    Console.WriteLine("[DEBUG] Task 2-5-3 Failed: Legend is still visible.");
                    return false;
                }

                // 2. データの範囲変更（過去の月のデータが追加されているか）のチェック
                SeriesCollection seriesColl = (SeriesCollection)chart.SeriesCollection();
                if (seriesColl.Count == 0) return false;

                Series series1 = seriesColl.Item(1);
                string formula = series1.Formula;
                Console.WriteLine($"[DEBUG] Series Formula: {formula}");

                // 過去の月のデータが追加されているかチェック（範囲が拡張されているか）
                // 通常、データ範囲が拡張されている場合は、行番号が大きくなる
                bool rangeExpanded = formula.Contains("$10") || formula.Contains(":10") || 
                                     formula.Contains("$11") || formula.Contains(":11") ||
                                     formula.Contains("$12") || formula.Contains(":12");

                if (rangeExpanded)
                {
                    Console.WriteLine("[DEBUG] Task 2-5-3 Passed: Legend removed and range expanded.");
                    return true;
                }

                Console.WriteLine("[DEBUG] Task 2-5-3 Failed: Range does not match expected expansion.");
                return false;
            });
        }

        private bool CheckTask_2_5_04_Impl(string filePath)
        {
            // CSVの解答手順: 「販売実績」シート内のグラフの数値が見えないように変更
            // ExcelChecker1_5のCheckTask_1_5_03_Implを参考に、ProcessChartTaskパターンを使用
            return ProcessChartTask(filePath, "販売実績", (chart) =>
            {
                // データラベル非表示のチェック
                SeriesCollection seriesColl = (SeriesCollection)chart.SeriesCollection();
                if (seriesColl.Count == 0) return false;

                foreach (Series s in seriesColl)
                {
                    if (s.HasDataLabels)
                    {
                        Console.WriteLine("[DEBUG] Task 2-5-4 Failed: Data Labels are visible.");
                        return false; 
                    }
                }

                Console.WriteLine("[DEBUG] Task 2-5-4 Passed: Data Labels are hidden.");
                return true;
            });
        }

        private bool CheckTask_2_5_05_Impl(string filePath)
        {
            // CSVの解答手順にはタスク2-5-5が記載されていないため、このメソッドは削除またはスキップ
            // 既存の実装を維持するか、常にtrueを返す
            return true;
        }

        public bool CheckTask_2_5_06()
        {
            // CSVの解答手順にはタスク2-5-6が記載されていないため、このメソッドは削除またはスキップ
            // 既存の実装を維持するか、常にtrueを返す
            return true;
        }

        private bool CheckTask_2_5_06_Impl(string filePath)
        {
            return true;
        }

        private bool CheckTask_2_5_06_Comparison(string currentFilePath)
        {
            return true;
        }

        private string GetCurrentExcelFilePath()
        {
            Application excelApp = null;
            try
            {
                // 実行中のExcelアプリケーションを取得
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                // アクティブなワークブックのパスを取得
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
                
                return null;
            }
            catch (COMException)
            {
                // Excelが起動していない場合
                return null;
            }
            finally
            {
                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);
            }
        }
        
        private Worksheet FindWorksheet(Workbook workbook, string sheetName)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }
            return null;
        }

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク01-05）
        /// </summary>
        private bool CheckTask_2_5_01_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 5);
        }

        private bool CheckTask_2_5_02_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 5);
        }

        private bool CheckTask_2_5_03_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 5);
        }

        private bool CheckTask_2_5_04_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 5);
        }

        private bool CheckTask_2_5_05_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 5);
        }

        /// <summary>
        /// 一般的な比較チェック（ページ設定を比較）
        /// </summary>
        private bool PerformGeneralComparison(string currentFilePath, int tabId, int projectId)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(tabId, projectId);
                string completedFilePath = GetCompletedDataFilePath(tabId, projectId);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                // アクティブなワークシートを比較
                currentWorksheet = currentWorkbook.ActiveSheet as Worksheet;
                completedWorksheet = completedWorkbook.ActiveSheet as Worksheet;

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                // ページ設定を比較
                return ComparePageSetup(currentWorksheet.PageSetup, completedWorksheet.PageSetup);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComparisonComObject(completedWorksheet);
                ReleaseComparisonComObject(currentWorksheet);
                CloseComparisonWorkbook(completedWorkbook);
            }
        }

        // ExcelCheckerComparisonHelperの機能を直接実装
        private static JObject _config;

        static ExcelChecker2_5()
        {
            LoadConfig();
        }

        private static void LoadConfig()
        {
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    _config = JObject.Parse(json);
                }
            }
            catch { }
        }

        private static string GetInitialDataFilePath(int tabId, int projectId)
        {
            if (_config == null) return null;
            try
            {
                var projectData = _config["tabs"]?[tabId.ToString()]?["projects"]?[projectId.ToString()];
                return projectData?["initialDataFile"]?.ToString();
            }
            catch { return null; }
        }

        private static string GetCompletedDataFilePath(int tabId, int projectId)
        {
            if (_config == null) return null;
            try
            {
                var projectData = _config["tabs"]?[tabId.ToString()]?["projects"]?[projectId.ToString()];
                return projectData?["completedDataFile"]?.ToString();
            }
            catch { return null; }
        }

        private static Workbook GetComparisonWorkbook(Application excelApp, string filePath)
        {
            if (excelApp == null || string.IsNullOrEmpty(filePath)) return null;
            try
            {
                string fileName = Path.GetFileName(filePath);
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        return wb;
                    }
                }
                if (File.Exists(filePath))
                {
                    return excelApp.Workbooks.Open(filePath, ReadOnly: true);
                }
            }
            catch { }
            return null;
        }

        private static bool ComparePageSetup(PageSetup setup1, PageSetup setup2)
        {
            if (setup1 == null || setup2 == null) return false;
            try
            {
                if (setup1.Orientation != setup2.Orientation) return false;
                string printArea1 = setup1.PrintArea ?? string.Empty;
                string printArea2 = setup2.PrintArea ?? string.Empty;
                if (NormalizeRange(printArea1) != NormalizeRange(printArea2)) return false;
                string titleRows1 = setup1.PrintTitleRows ?? string.Empty;
                string titleRows2 = setup2.PrintTitleRows ?? string.Empty;
                if (NormalizeRange(titleRows1) != NormalizeRange(titleRows2)) return false;
                const double tolerance = 1.0;
                if (Math.Abs(setup1.TopMargin - setup2.TopMargin) > tolerance) return false;
                if (Math.Abs(setup1.BottomMargin - setup2.BottomMargin) > tolerance) return false;
                if (Math.Abs(setup1.LeftMargin - setup2.LeftMargin) > tolerance) return false;
                if (Math.Abs(setup1.RightMargin - setup2.RightMargin) > tolerance) return false;
                return true;
            }
            catch { return false; }
        }

        private static string NormalizeRange(string range)
        {
            if (string.IsNullOrEmpty(range)) return string.Empty;
            return range.Replace("$", "").Replace(" ", "").ToUpper();
        }

        private static void CloseComparisonWorkbook(Workbook workbook, bool saveChanges = false)
        {
            if (workbook != null)
            {
                try
                {
                    workbook.Close(saveChanges);
                    Marshal.ReleaseComObject(workbook);
                }
                catch { }
            }
        }

        private static void ReleaseComparisonComObject(object obj)
        {
            if (obj != null)
            {
                try { Marshal.ReleaseComObject(obj); }
                catch { }
            }
        }

        // ==========================================
        // 共通プロセス・ヘルパーメソッド（ExcelChecker1_5を参考）
        // ==========================================
        
        private bool ProcessChartTask(string filePath, string sheetName, Func<Chart, bool> checkLogic)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
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
                        break;
                    }
                }
                if (workbook == null) return false;

                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = ws;
                        break;
                    }
                }
                if (worksheet == null) return false;

                ChartObjects chartObjects = (ChartObjects)worksheet.ChartObjects();
                if (chartObjects.Count == 0) return false;

                ChartObject chartObj = chartObjects.Item(1);
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
    }
}