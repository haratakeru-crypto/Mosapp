using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace Libraries.Group2
{
    public class ExcelChecker2_9
    {
        public bool CheckTask_2_9_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_9_01_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_9_01_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_9_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_9_02_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_9_02_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_9_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_9_03_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_9_03_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_9_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_9_04_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_9_04_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_9_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_9_05_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_9_05_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_9_01_Impl(string filePath)
        {
            // CSVの解答手順: 「担当者マスター」シートのE列に勤続年数が5より大きければ「あり」、そうでなければ「なし」と表示
            // ExcelChecker1_10のCheckTask_1_10_01_Implを参考に、IF関数をチェック
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

                worksheet = FindWorksheet(workbook, "担当者マスター");
                if (worksheet == null) return false;

                // E5セルのIF関数をチェック
                Range targetCell = worksheet.Range["E5"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // IF関数とD5>5の条件、「あり」「なし」の値をチェック
                    return normalizedFormula.Contains("IF(") &&
                           normalizedFormula.Contains("D5>5") &&
                           normalizedFormula.Contains("あり") &&
                           normalizedFormula.Contains("なし") &&
                           !normalizedFormula.Contains("$");
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

        private bool CheckTask_2_9_02_Impl(string filePath)
        {
            // CSVの解答手順: 「出張精算」シートの「手当金額」の列に距離が300㎞以上であれば「10000」、なければ「5000」と表示
            // ExcelChecker1_10のCheckTask_1_10_02_Implを参考に、IF関数をチェック
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

                worksheet = FindWorksheet(workbook, "出張精算");
                if (worksheet == null) return false;

                // G8セルのIF関数をチェック
                Range targetCell = worksheet.Range["G8"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // IF関数とF8>=300の条件、10000、5000の値をチェック
                    return normalizedFormula.Contains("IF(") &&
                           normalizedFormula.Contains("F8>=300") &&
                           normalizedFormula.Contains("10000") &&
                           normalizedFormula.Contains("5000") &&
                           !normalizedFormula.Contains("$");
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

        private bool CheckTask_2_9_03_Impl(string filePath)
        {
            // CSVの解答手順: 「売上一覧」シートの在庫が13％以下であれば「在庫を補充」、そうでなければ空欄が補充の有無の列に表示
            // ExcelChecker1_10のCheckTask_1_10_03_Implを参考に、IF関数をチェック
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

                // 補充の有無の列でIF関数をチェック（通常G7など）
                Range targetCell = worksheet.Range["G7"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // IF関数と在庫<=13%または<=0.13の条件、「在庫を補充」の値、空文字列をチェック
                    bool hasIF = normalizedFormula.Contains("IF(");
                    bool hasCondition = normalizedFormula.Contains("<=13%") || normalizedFormula.Contains("<=0.13");
                    bool hasMessage = normalizedFormula.Contains("在庫を補充");
                    bool hasEmptyString = normalizedFormula.Contains("\"\"");
                    bool noAbsoluteRef = !normalizedFormula.Contains("$");
                    
                    return hasIF && hasCondition && hasMessage && hasEmptyString && noAbsoluteRef;
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

        private bool CheckTask_2_9_04_Impl(string filePath)
        {
            // CSVの解答手順: 「模試結果」シート内の表の通し番号を関数を使って1から自動で順に表示
            // ExcelChecker1_10のCheckTask_1_10_04_Implを参考に、SEQUENCE関数をチェック
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

                // 通し番号の列でSEQUENCE関数をチェック（通常B8など）
                Range targetCell = worksheet.Range["B8"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // SEQUENCE関数と22行、1列、開始値1、目盛り1をチェック
                    return normalizedFormula.Contains("SEQUENCE(") &&
                           normalizedFormula.Contains("22") &&
                           normalizedFormula.Contains("1,1,1");
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

        private bool CheckTask_2_9_05_Impl(string filePath)
        {
            // CSVの解答手順: 「営業予定」シートの時間変更の列に9時から15分ごとに内容が終わるように変更
            // ExcelChecker1_10のCheckTask_1_10_05_Implを参考に、SEQUENCE関数をチェック
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

                worksheet = FindWorksheet(workbook, "営業予定");
                if (worksheet == null) return false;

                // 時間変更の列でSEQUENCE関数をチェック（通常D10など）
                Range targetCell = worksheet.Range["D10"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // SEQUENCE関数と目盛り0.25（15分=0.25時間）をチェック
                    return normalizedFormula.Contains("SEQUENCE(") &&
                           normalizedFormula.Contains("0.25");
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

        public bool CheckTask_2_9_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                bool existingCheck = CheckTask_2_9_06_Impl(filePath);
                bool comparisonCheck = CheckTask_2_9_06_Comparison(filePath);
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_9_06_Impl(string filePath)
        {
            // CSVの解答手順: 「担当リスト」シートの売上順担当者名の列に売上が高い順に並び替えて表示
            // ExcelChecker1_10のCheckTask_1_10_06_Implを参考に、SORT関数をチェック
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

                worksheet = FindWorksheet(workbook, "担当リスト");
                if (worksheet == null) return false;

                // 売上順担当者名の列でSORT関数をチェック（通常D10など）
                Range targetCell = worksheet.Range["D10"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // SORT関数と並べ替えインデックス2、並べ替え順序-1（降順）をチェック
                    return normalizedFormula.Contains("SORT(") &&
                           normalizedFormula.Contains("A10:B15") &&
                           normalizedFormula.Contains("2") &&
                           normalizedFormula.Contains("-1");
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

        public bool CheckTask_2_9_07()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                bool existingCheck = CheckTask_2_9_07_Impl(filePath);
                bool comparisonCheck = CheckTask_2_9_07_Comparison(filePath);
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_9_07_Impl(string filePath)
        {
            // CSVの解答手順: 「売上一覧」シートの税込価格の列に税込み金額を求める、税込み価格は「単価×税率」で計算、税率はセルを参照
            // ExcelChecker1_10のCheckTask_1_10_07_Implを参考に、数式をチェック
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

                // 税込価格の列で数式をチェック（通常J7など）
                Range targetCell = worksheet.Range["J7"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // F7*$L$11の絶対参照をチェック（税率はセル参照）
                    return normalizedFormula.Contains("F7") &&
                           normalizedFormula.Contains("$L$11") &&
                           normalizedFormula.Contains("*");
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

        private bool CheckTask_2_9_06_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 9);
        }

        private bool CheckTask_2_9_07_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 9);
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
            }
            catch (Exception)
            {
                // Excel アプリケーションが見つからない場合
            }
            return string.Empty;
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
        private bool CheckTask_2_9_01_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 9);
        }

        private bool CheckTask_2_9_02_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 9);
        }

        private bool CheckTask_2_9_03_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 9);
        }

        private bool CheckTask_2_9_04_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 9);
        }

        private bool CheckTask_2_9_05_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 9);
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

        static ExcelChecker2_9()
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
        // 共通プロセス・ヘルパーメソッド（ExcelChecker1_9を参考）
        // ==========================================
        
        private bool ProcessSheet(string filePath, string sheetName, Func<Worksheet, bool> checkLogic)
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

                return checkLogic(worksheet);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error in ProcessSheet: {ex.Message}");
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }
    }
}