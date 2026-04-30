using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace Libraries.Group2
{
    public class ExcelChecker2_7
    {
        public bool CheckTask_2_7_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_7_01_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_7_01_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_7_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_7_02_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_7_02_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_7_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_7_03_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_7_03_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_7_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_7_04_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_7_04_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_7_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_7_05_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_7_05_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_7_01_Impl(string filePath)
        {
            // CSVの解答手順: 「文化祭」シートの4日目の売上合計の最高金額をJ6に表示
            // ExcelChecker1_7のCheckTask_1_7_02_Implを参考に、MAX関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "文化祭");
                if (worksheet == null) return false;
                
                // J6セルのMAX関数をチェック
                Range targetCell = worksheet.Range["J6"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    
                    // MAX関数と4日目の売上合計の範囲をチェック（テーブル参照も許容）
                    bool result = normalizedFormula.Contains("MAX(") && 
                                  (normalizedFormula.Contains("G8:G16") ||           // 通常の範囲参照
                                   normalizedFormula.Contains("文化祭") ||            // テーブル名
                                   normalizedFormula.Contains("[[4日目]:[4日目]]")) &&  // テーブル列範囲
                                  !normalizedFormula.Contains("$");
                    
                    return result;
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

        private bool CheckTask_2_7_02_Impl(string filePath)
        {
            // CSVの解答手順: 「文化祭」シートのテーブルのH列の式を下まで完成
            // ExcelChecker1_7のCheckTask_1_7_01_Implを参考に、数式のコピーをチェック
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
                
                worksheet = FindWorksheet(workbook, "文化祭");
                if (worksheet == null) return false;
                
                // H8からH16までの範囲で数式がコピーされているかチェック
                Range h8Cell = worksheet.Range["H8"];
                Range targetRange = worksheet.Range["H9:H16"];
                
                bool h8HasFormula = h8Cell.HasFormula is bool && (bool)h8Cell.HasFormula;
                if (h8HasFormula)
                {
                    foreach (Range cell in targetRange.Cells)
                    {
                        bool cellHasFormula = cell.HasFormula is bool && (bool)cell.HasFormula;
                        if (!cellHasFormula)
                        {
                            return false; // 一つでも数式がないセルがあればfalse
                        }
                    }
                    return true;
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

        private bool CheckTask_2_7_03_Impl(string filePath)
        {
            // CSVの解答手順: 「試験結果」シートのE9に合計点がある受験生の数を表示
            // ExcelChecker1_7のCheckTask_1_7_03_Implを参考に、COUNT関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "試験結果");
                if (worksheet == null) return false;
                
                // E9セルのCOUNT関数をチェック
                Range targetCell = worksheet.Range["E9"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // COUNT関数と適切な範囲の組み合わせをチェック（テーブル参照も許容）
                    return normalizedFormula.Contains("COUNT(") && 
                           (normalizedFormula.Contains("合計点") ||               // テーブル列名
                            normalizedFormula.Contains("試験結果") ||               // テーブル名
                            normalizedFormula.Contains("[[合計点]:[合計点]]")) &&   // テーブル列範囲
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

        private bool CheckTask_2_7_04_Impl(string filePath)
        {
            // CSVの解答手順: 「試験結果」シートのF9に合計点のない受験生の数をカウント
            // ExcelChecker1_7のCheckTask_1_7_04_Implを参考に、COUNTBLANK関数をチェック
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
                
                worksheet = FindWorksheet(workbook, "試験結果");
                if (worksheet == null) return false;
                
                // F9セルのCOUNTBLANK関数をチェック
                Range targetCell = worksheet.Range["F9"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // COUNTBLANK関数と適切な範囲の組み合わせをチェック（テーブル参照も許容）
                    return normalizedFormula.Contains("COUNTBLANK(") && 
                           (normalizedFormula.Contains("合計点") ||               // テーブル列名
                            normalizedFormula.Contains("試験結果") ||               // テーブル名
                            normalizedFormula.Contains("[[合計点]:[合計点]]")) &&   // テーブル列範囲
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

        private bool CheckTask_2_7_05_Impl(string filePath)
        {
            // CSVの解答手順: RANDBETWEEN関数を使って「模試結果」シート内の受験番号に1～22の番号を入力
            // ExcelChecker1_7のCheckTask_1_7_05_Implを参考に、RANDBETWEEN関数をチェック
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
                
                // 受験番号の列でRANDBETWEEN関数をチェック（通常B8からB29など）
                Range targetCell = worksheet.Range["B8"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // RANDBETWEEN関数とパラメータ1,22の組み合わせをチェック
                    return normalizedFormula.Contains("RANDBETWEEN(") && 
                           normalizedFormula.Contains("1,22");
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

        public bool CheckTask_2_7_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                bool existingCheck = CheckTask_2_7_06_Impl(filePath);
                bool comparisonCheck = CheckTask_2_7_06_Comparison(filePath);
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_7_06_Impl(string filePath)
        {
            // CSVの解答手順: 「模試結果」シート内の学籍番号の最初の2文字が学科のアルファベットに表示されるように
            // ExcelChecker1_7のCheckTask_1_7_06_Implを参考に、LEFT関数をチェック
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
                
                // 学科の列でLEFT関数をチェック（通常C8など）
                Range targetCell = worksheet.Range["C8"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // LEFT関数と相対参照パラメータ（学籍番号の列,2）の組み合わせをチェック
                    return normalizedFormula.Contains("LEFT(") && 
                           normalizedFormula.Contains(",2") &&
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

        public bool CheckTask_2_7_07()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                bool existingCheck = CheckTask_2_7_07_Impl(filePath);
                bool comparisonCheck = CheckTask_2_7_07_Comparison(filePath);
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_7_07_Impl(string filePath)
        {
            // CSVの解答手順: 「売上一覧」シートの商品の種類の列に商品名を重複せずにすべて表示
            // ExcelChecker1_7のCheckTask_1_7_07_Implを参考に、UNIQUE関数をチェック
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
                
                // 商品の種類の列でUNIQUE関数をチェック（通常I5など）
                Range targetCell = worksheet.Range["I5"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    
                    // UNIQUE関数と適切な範囲の組み合わせをチェック（テーブル参照も許容）
                    bool result = normalizedFormula.Contains("UNIQUE(") && 
                                  (normalizedFormula.Contains("商品名") ||           // テーブル列名
                                   normalizedFormula.Contains("売上一覧") ||            // テーブル名
                                   normalizedFormula.Contains("[[商品名]:[商品名]]")) &&   // テーブル列範囲
                                  !normalizedFormula.Contains("$");
                    
                    return result;
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

        private bool CheckTask_2_7_06_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 7);
        }

        private bool CheckTask_2_7_07_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 7);
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
        private bool CheckTask_2_7_01_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 7);
        }

        private bool CheckTask_2_7_02_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 7);
        }

        private bool CheckTask_2_7_03_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 7);
        }

        private bool CheckTask_2_7_04_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 7);
        }

        private bool CheckTask_2_7_05_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 7);
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

        static ExcelChecker2_7()
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
    }
}