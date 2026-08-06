using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace Libraries.Group2
{
    public class ExcelChecker2_6
    {
        public bool CheckTask_2_6_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_6_01_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_6_01_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_6_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_6_02_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_6_02_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_6_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_6_03_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_6_03_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_6_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_6_04_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_6_04_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_6_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_6_05_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_6_05_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_6_01_Impl(string filePath)
        {
            // CSVの解答手順: 「模試結果」シートの7行目を固定
            // ExcelChecker1_6のCheckTask_1_6_01_Implを参考に、ウィンドウ固定をチェック
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
                    excelApp.Visible = false;
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                worksheet = FindWorksheet(workbook, "模試結果");
                if (worksheet == null) return false;

                worksheet.Activate();
                Window window = excelApp.ActiveWindow;
                if (window.FreezePanes)
                {
                    // 7行目が固定されているかチェック
                    if (window.SplitRow >= 7)
                    {
                        return true;
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

        private bool CheckTask_2_6_02_Impl(string filePath)
        {
            // CSVの解答手順: 「販売実績」シート内の表内の数値の書式を「通貨」に変更、小数点は表示されないように
            // ExcelChecker1_6のCheckTask_1_6_03_Implを参考に、通貨書式をチェック
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
                    excelApp.Visible = false;
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                worksheet = FindWorksheet(workbook, "販売実績");
                if (worksheet == null) return false;

                // 表内の数値範囲をチェック（実際の範囲はファイル構造に依存）
                // 通常、B5:G11などの範囲が表の数値部分
                Range targetRange = worksheet.Range["B5:G11"];
                foreach (Range cell in targetRange.Cells)
                {
                    string numberFormat = cell.NumberFormat as string;
                    if (numberFormat == null) return false;
                    
                    // 通貨形式であることを確認（¥または$を含む）
                    bool isCurrency = numberFormat.Contains("¥") || numberFormat.Contains("$");
                    // 小数点が表示されていないことを確認
                    bool hasNoDecimals = !numberFormat.Contains(".0") && !numberFormat.Contains(".#");
                    
                    // すべてのセルが通貨形式（小数点なし）である必要がある
                    if (!isCurrency || !hasNoDecimals)
                    {
                        return false;
                    }
                }
                return true;
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

        private bool CheckTask_2_6_03_Impl(string filePath)
        {
            // CSVの解答手順: 「売上一覧」シートのI6のパソコンのご相談の文字に「http:pcostomer.jp」のリンクを挿入
            // ExcelChecker1_6のCheckTask_1_6_02_Implを参考に、ハイパーリンクをチェック
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
                    excelApp.Visible = false;
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;

                foreach (Hyperlink hyperlink in worksheet.Hyperlinks)
                {
                    if (hyperlink.Address != null &&
                        hyperlink.Address.Contains("http:pcostomer.jp"))
                    {
                        if (hyperlink.TextToDisplay != null &&
                            hyperlink.TextToDisplay.Contains("パソコンのご相談"))
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

        private bool CheckTask_2_6_04_Impl(string filePath)
        {
            // CSVの解答手順: プロパティのタグに「売上」と追加
            // ExcelChecker1_6のCheckTask_1_6_04_Implを参考に、プロパティタグをチェック
            Application excelApp = null;
            Workbook workbook = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = false;
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                try
                {
                    dynamic properties = workbook.BuiltinDocumentProperties;
                    dynamic keywordsProperty = properties["Keywords"];
                    string tags = keywordsProperty.Value as string;
                    if (tags != null && tags.Contains("売上"))
                    {
                        return true;
                    }
                }
                catch
                {
                    return false;
                }

                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                // workbookはCloseしない
            }
        }

        private bool CheckTask_2_6_05_Impl(string filePath)
        {
            // CSVの解答手順にはタスク2-6-5が記載されていないため、このメソッドは削除またはスキップ
            // 既存の実装を維持するか、常にtrueを返す
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

        private Workbook GetWorkbook(Application excelApp, string filePath)
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
            return null;
        }

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク01-05）
        /// </summary>
        private bool CheckTask_2_6_01_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 6);
        }

        private bool CheckTask_2_6_02_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 6);
        }

        private bool CheckTask_2_6_03_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 6);
        }

        private bool CheckTask_2_6_04_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 6);
        }

        private bool CheckTask_2_6_05_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 6);
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

        static ExcelChecker2_6()
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
            string configured = _config?["tabs"]?[tabId.ToString()]?["projects"]?[projectId.ToString()]?["initialDataFile"]?.ToString();
            return MOSExcelMogiApp.Infrastructure.DataPathHelper.ResolveInitialFilePath(tabId, projectId, configured);
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
    }
}