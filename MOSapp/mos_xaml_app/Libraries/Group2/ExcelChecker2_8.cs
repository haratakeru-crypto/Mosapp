using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace Libraries.Group2
{
    public class ExcelChecker2_8
    {
        public bool CheckTask_2_8_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_8_01_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_8_01_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_8_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_8_02_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_8_02_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_8_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_8_03_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_8_03_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_8_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_8_04_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_8_04_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_8_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // タスク8-5はメールアドレス列のCONCATのみが対象。ページ設定の比較は不要なためImplのみで判定する
                return CheckTask_2_8_05_Impl(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_8_01_Impl(string filePath)
        {
            // CSVの解答手順: キャンパス別試験結果シートのB6:C26に「学生」という名前を付ける
            // ExcelChecker1_8のCheckTask_1_8_01_Implを参考に、名前定義をチェック
            Application excelApp = null;
            Workbook workbook = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                // 名前「学生」を探す（スコープの違いを考慮して EndsWith を使用）
                foreach (Name name in workbook.Names)
                {
                    try
                    {
                        // "学生" または "キャンパス別試験結果!学生" などにマッチさせる
                        if (name.Name == "学生" || name.Name.EndsWith("!学生"))
                        {
                            Range range = name.RefersToRange;
                            if (range != null)
                            {
                                // 参照先シートが「キャンパス別試験結果」であること
                                if (range.Worksheet.Name == "キャンパス別試験結果")
                                {
                                    // 参照先アドレスが B6:C26 であること
                                    string address = range.Address.Replace("$", "");
                                    if (address == "B6:C26")
                                    {
                                        return true;
                                    }
                                }
                            }
                        }
                    }
                    catch { continue; }
                }
                return false;
            }
            catch { return false; }
        }

        private bool CheckTask_2_8_02_Impl(string filePath)
        {
            // CSVの解答手順: 「消費税」という名前付き範囲に移動し、数字を「10％」と変更
            // ExcelChecker1_8のCheckTask_1_8_02_Implを参考に、名前範囲の値をチェック
            Application excelApp = null;
            Workbook workbook = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                foreach (Name name in workbook.Names)
                {
                    try
                    {
                        if (name.Name == "消費税" || name.Name.EndsWith("!消費税"))
                        {
                            Range range = name.RefersToRange;
                            if (range != null)
                            {
                                string value = range.Text.ToString();
                                if (!string.IsNullOrEmpty(value) && value.Contains("10%"))
                                {
                                    return true;
                                }
                            }
                        }
                    }
                    catch { continue; }
                }
                return false;
            }
            catch { return false; }
        }

        private bool CheckTask_2_8_03_Impl(string filePath)
        {
            // CSVの解答手順: 「文化祭」シートのJ7に全部活の金額を合計した数値を「各売上合計」を使い表示、セル参照や数値は使わず、名前付き範囲を使う
            // ExcelChecker1_8のCheckTask_1_8_03_Implを参考に、SUM関数と名前定義をチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "文化祭");
                if (worksheet == null) return false;
                
                // 模試②用：解答手順あり模擬試験①問題文.csvを参照（J8に表示）
                Range targetCell = worksheet.Range["J8"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalized = formula.Replace(" ", "").ToUpper();

                    // 1. SUM(各売上合計) が含まれているか（必須）
                    bool hasName = normalized.Contains("SUM(各売上合計)") || normalized.Contains("=SUM(各売上合計)");
                    
                    if (!hasName) return false;

                    // 2. セル番地（数字）が含まれていないかチェック
                    // 名前定義「各売上合計」を使っていれば、数式中に "7" や "16" といった行番号は出ないはず
                    if (normalized.Contains(":"))
                    {
                        // SUM(各売上合計) があるのに : があるのはおかしい
                        return false;
                    }

                    return true;
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_2_8_04_Impl(string filePath)
        {
            // CSVの解答手順: キャンパス別試験結果シートの表内の学部学科の列に学部と学科を付け加えて、表を完成
            // ExcelChecker1_8のCheckTask_1_8_04_Implを参考に、CONCAT関数をチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "キャンパス別試験結果");
                if (worksheet == null) return false;
                
                // 学部学科の列でCONCAT関数をチェック（通常G7など）
                Range targetCell = worksheet.Range["G7"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalized = formula.Replace(" ", "").ToUpper();
                    // CONCAT関数とE7:F7の範囲をチェック
                    return normalized.Contains("CONCAT(E7:F7)") || normalized.Contains("=CONCAT(E7:F7)");
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_2_8_05_Impl(string filePath)
        {
            // CSVの解答手順: 「担当者マスター」シートの表のアカウントとアドレス「@rabbit.ac.jp」を組み合わせてメールアドレスの列に表示
            // ExcelChecker1_8のCheckTask_1_8_05_Implを参考に、CONCAT関数をチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "担当者マスター");
                if (worksheet == null) return false;
                
                // メールアドレスの列でCONCAT関数をチェック（通常H5など）
                Range targetCell = worksheet.Range["H5"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalized = formula.Replace(" ", "").ToUpper();
                    // CONCAT関数とG5と"@RABBIT.AC.JP"の組み合わせをチェック（問題文は @rabbit.ac.jp）
                    return normalized.Contains("CONCAT(G5,\"@RABBIT.AC.JP\")") || normalized.Contains("=CONCAT(G5,\"@RABBIT.AC.JP\")");
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        public bool CheckTask_2_8_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                bool existingCheck = CheckTask_2_8_06_Impl(filePath);
                bool comparisonCheck = CheckTask_2_8_06_Comparison(filePath);
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_8_06_Impl(string filePath)
        {
            // CSVの解答手順: 「売上一覧」シートのタグ列に商品型番-商品名と表示
            // ExcelChecker1_8のCheckTask_1_8_06_Implを参考に、CONCAT関数をチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;
                
                // タグ列でCONCAT関数をチェック（通常J7など）
                Range targetCell = worksheet.Range["J7"];
                string formula = targetCell.Formula as string;
                
                if (formula != null)
                {
                    string normalized = formula.Replace(" ", "").ToUpper();
                    // CONCAT関数と商品型番-商品名の組み合わせをチェック
                    // 例: CONCAT(A7,"-",B7) または CONCAT(A7,"-",商品名)
                    return normalized.Contains("CONCAT(") && 
                           normalized.Contains("-") &&
                           !normalized.Contains("$");
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_2_8_06_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 8);
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
        private bool CheckTask_2_8_01_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 8);
        }

        private bool CheckTask_2_8_02_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 8);
        }

        private bool CheckTask_2_8_03_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 8);
        }

        private bool CheckTask_2_8_04_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 8);
        }

        private bool CheckTask_2_8_05_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 8);
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

        static ExcelChecker2_8()
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