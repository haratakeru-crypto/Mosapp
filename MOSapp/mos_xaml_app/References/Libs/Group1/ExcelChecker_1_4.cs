using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelChecker4
{
    public class ExcelChecker4
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\project4.xlsx";

        /// <summary>
        /// Project4のタスク4-1をチェックする
        /// シート[上半期売上]の「売上状況」の列に、1月から6月の売上の大小を表す縦棒スパークラインを挿入
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_04_Task_04_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_04_Task_04_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project4のタスク4-2をチェックする
        /// シート[５年間売上]の表のデータをもとに、商品ごとに2021年度と2022年度の売上を比較する積み上げ縦棒グラフを作成
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_04_Task_04_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_04_Task_04_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project4のタスク4-3をチェックする
        /// シート[下半期売上]の表のデータをもとに、売上の合計を表す3-D円グラフを作成
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_04_Task_04_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_04_Task_04_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project4のタスク4-4をチェックする
        /// シート[商品別売上]のグラフに、代替テキスト「有楽町店の売上グラフ」を追加
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_04_Task_04_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_04_Task_04_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_4()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: project4.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: project4.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_04_Task_04_01(TARGET_FILE_PATH);
                results.Add($"Task 4-1 (縦棒スパークライン): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_04_Task_04_02(TARGET_FILE_PATH);
                results.Add($"Task 4-2 (積み上げ縦棒グラフ): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_04_Task_04_03(TARGET_FILE_PATH);
                results.Add($"Task 4-3 (3-D円グラフ): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_04_Task_04_04(TARGET_FILE_PATH);
                results.Add($"Task 4-4 (代替テキスト): {(task4 ? "OK" : "NG")}");
                
                return string.Join("\n", results);
            }
            catch (Exception ex)
            {
                return $"エラー: {ex.Message}";
            }
        }

        private bool IsExcelFileOpen(string filePath)
        {
            Application excelApp = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                foreach (Workbook workbook in excelApp.Workbooks)
                {
                    if (string.Equals(workbook.FullName, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);
            }
        }

        private string GetCurrentExcelFilePath()
        {
            Application excelApp = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
                return null;
            }
            catch (COMException)
            {
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

        private Workbook GetWorkbook(Application excelApp, string filePath)
        {
            string fileName = System.IO.Path.GetFileName(filePath);
            foreach (Workbook wb in excelApp.Workbooks)
            {
                if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                    wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                {
                    return wb;
                }
            }
            return null;
        }

        private bool CheckProject_04_Task_04_01(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "上半期売上");
                if (worksheet == null) return false;
                
                // セル範囲I5:I10にスパークラインが存在するかチェック
                Range targetRange = worksheet.Range["I5:I10"];
                foreach (Range cell in targetRange.Cells)
                {
                    if (cell.SparklineGroups.Count > 0)
                    {
                        // スパークラインの種類が縦棒かチェック
                        var sparklineGroup = cell.SparklineGroups[1];
                        if (sparklineGroup.Type == XlSparkType.xlSparkColumn)
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckProject_04_Task_04_02(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "５年間売上");
                if (worksheet == null) return false;
                
                // シートにグラフが存在するかチェック
                ChartObjects chartObjects = (ChartObjects)worksheet.ChartObjects();
                if (chartObjects.Count > 0)
                {
                    foreach (ChartObject chartObj in chartObjects)
                    {
                        Chart chart = chartObj.Chart;
                        // 積み上げ縦棒グラフかチェック
                        if (chart.ChartType == XlChartType.xlColumnStacked)
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckProject_04_Task_04_03(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "下半期売上");
                if (worksheet == null) return false;
                
                // シートにグラフが存在するかチェック
                ChartObjects chartObjects = (ChartObjects)worksheet.ChartObjects();
                if (chartObjects.Count > 0)
                {
                    foreach (ChartObject chartObj in chartObjects)
                    {
                        Chart chart = chartObj.Chart;
                        // 3-D円グラフかチェック
                        if (chart.ChartType == XlChartType.xlPie3D)
                        {
                            // 凡例が表示されているかチェック
                            if (chart.HasLegend)
                            {
                                return true;
                            }
                        }
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckProject_04_Task_04_04(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "商品別売上");
                if (worksheet == null) return false;
                
                // シートにグラフが存在するかチェック
                ChartObjects chartObjects = (ChartObjects)worksheet.ChartObjects();
                if (chartObjects.Count > 0)
                {
                    foreach (ChartObject chartObj in chartObjects)
                    {
                        Chart chart = chartObj.Chart;
                        // 代替テキストをチェック
                        try
                        {
                            string altText = chart.AlternativeText;
                            if (!string.IsNullOrEmpty(altText) && altText.Contains("有楽町店の売上グラフ"))
                            {
                                return true;
                            }
                        }
                        catch
                        {
                            // 代替テキストが取得できない場合は継続
                            continue;
                        }
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }
    }
} 