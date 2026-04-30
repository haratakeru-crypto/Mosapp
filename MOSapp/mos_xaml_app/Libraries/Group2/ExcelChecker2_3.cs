using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace Libraries.Group2
{
    public class ExcelChecker2_3
    {
        public bool CheckTask_2_3_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_3_01_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_3_01_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_3_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_3_02_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_3_02_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_3_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_3_03_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_3_03_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_3_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_3_04_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_3_04_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_3_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_3_05_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_3_05_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_3_01_Impl(string filePath)
        {
            // CSVの解答手順: 「販売実績」シートのセルA5を見出し3スタイルに
            // ExcelChecker1_3のCheckTask_1_3_01_Implを参考に、セルスタイルをチェック
            return CheckTaskBasic(filePath, "販売実績", (worksheet) =>
            {
                Range targetCell = worksheet.Range["A5"];
                dynamic style = targetCell.Style;
                string styleName = style.NameLocal; // "見出し3"
                string styleNameEng = style.Name;   // "Heading 3"
                return styleName == "見出し3" || styleName == "Heading 3" || 
                       styleNameEng == "見出し3" || styleNameEng == "Heading 3";
            });
        }

        private bool CheckTask_2_3_02_Impl(string filePath)
        {
            // CSVの解答手順: キャンパス別試験結果シートの氏名の列C7:C26に左インデントを2つ追加
            // ExcelChecker1_3のCheckTask_1_3_02_Implを参考に、左インデントをチェック
            return CheckTaskBasic(filePath, "キャンパス別試験結果", (ws) =>
            {
                Range targetRange = ws.Range["C7:C26"];

                // 範囲外チェック (C6:見出し, C27:下)
                Range topNeighbor = ws.Range["C6"];
                Range bottomNeighbor = ws.Range["C27"];

                int topIndent = Convert.ToInt32(topNeighbor.IndentLevel);
                int bottomIndent = Convert.ToInt32(bottomNeighbor.IndentLevel);

                if (topIndent != 0)
                {
                    Console.WriteLine("[DEBUG] Task 2-3-2 Failed: Header C6 is indented.");
                    return false;
                }
                if (bottomIndent != 0)
                {
                    Console.WriteLine("[DEBUG] Task 2-3-2 Failed: Bottom cell C27 is indented.");
                    return false;
                }

                // 本体チェック
                foreach (Range cell in targetRange)
                {
                    int indent = Convert.ToInt32(cell.IndentLevel);
                    if (indent != 2) return false;
                }
                return true;
            });
        }

        private bool CheckTask_2_3_03_Impl(string filePath)
        {
            // CSVの解答手順: キャンパス別試験結果シートのセルB4内の文字をB4とC4の中央に配置
            // ExcelChecker1_3のCheckTask_1_3_03_Implを参考に、選択範囲内で中央をチェック
            return CheckTaskBasic(filePath, "キャンパス別試験結果", (ws) =>
            {
                Range targetCell = ws.Range["B4"];
                
                const int XlHAlignCenterAcrossSelection = 7;

                int align = 0;
                try 
                { 
                    align = Convert.ToInt32(targetCell.HorizontalAlignment); 
                } 
                catch 
                { 
                    return false; 
                }

                if (align != XlHAlignCenterAcrossSelection) 
                {
                    Console.WriteLine($"[DEBUG] Task 2-3-3 Failed: Alignment is {align} (Expected 7).");
                    return false;
                }

                // 範囲外チェック (C4)
                Range rightNeighbor = ws.Range["C4"];
                int rightAlign = 0;
                try
                {
                    rightAlign = Convert.ToInt32(rightNeighbor.HorizontalAlignment);
                }
                catch { }

                if (rightAlign == XlHAlignCenterAcrossSelection)
                {
                    Console.WriteLine("[DEBUG] Task 2-3-3 Failed: Range extends too far (C4 is included).");
                    return false;
                }

                Console.WriteLine("[DEBUG] Task 2-3-3 Passed.");
                return true;
            });
        }

        private bool CheckTask_2_3_04_Impl(string filePath)
        {
            // CSVの解答手順: 担当者マスターシートの表にG11:J16をコピーして列幅を保持して貼り付け
            // ExcelChecker1_3のCheckTask_1_3_04_Implを参考に、列幅保持コピーをチェック
            // 模試②用：解答手順あり模擬試験①問題文.csvを参照
            return CheckTaskBasic(filePath, "担当者マスター", (worksheet) =>
            {
                Range sourceRange = worksheet.Range["G11:J16"];
                // 貼り付け先は通常、元の位置の左側（A11:D16など）に貼り付けられる
                // ExcelChecker1_3のパターンに合わせて、コピー元とコピー先の列幅を比較
                Range targetRange = worksheet.Range["A11:D16"];

                // 列幅チェック（コピー元とコピー先の列幅が一致していることを確認）
                for (int j = 1; j <= 4; j++)
                {
                    Range sourceCol = (Range)sourceRange.Cells[1, j];
                    Range targetCol = (Range)targetRange.Cells[1, j];
                    
                    double w1 = Convert.ToDouble(sourceCol.ColumnWidth);
                    double w2 = Convert.ToDouble(targetCol.ColumnWidth);

                    // 列幅の差が0.1より大きい場合は不合格
                    if (Math.Abs(w1 - w2) > 0.1) 
                    {
                        Console.WriteLine($"[DEBUG] Task 2-3-4 Failed: Column width mismatch at column {j}.");
                        return false;
                    }
                }
                Console.WriteLine("[DEBUG] Task 2-3-4 Passed: Column widths match.");
                return true;
            });
        }

        private bool CheckTask_2_3_05_Impl(string filePath)
        {
            // CSVの解答手順: 「営業予定」シートのC10：C15の時間に取り消し線を付け
            // ExcelChecker1_3のCheckTask_1_3_05_Implを参考に、取り消し線をチェック
            return CheckTaskBasic(filePath, "営業予定", (ws) =>
            {
                Range targetRange = ws.Range["C10:C15"];
                
                // 範囲外チェック (C9, C16)
                Range topNeighbor = ws.Range["C9"];
                Range bottomNeighbor = ws.Range["C16"];
                
                bool topStrike = topNeighbor.Font.Strikethrough is bool bTop && bTop;
                bool bottomStrike = bottomNeighbor.Font.Strikethrough is bool bBot && bBot;

                if (topStrike || bottomStrike)
                {
                    Console.WriteLine("[DEBUG] Task 2-3-5 Failed: Range incorrect.");
                    return false;
                }

                foreach (Range cell in targetRange)
                {
                    if (!(cell.Font.Strikethrough is bool b && b)) return false;
                }
                return true;
            });
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
        private bool CheckTask_2_3_01_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 3);
        }

        private bool CheckTask_2_3_02_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 3);
        }

        private bool CheckTask_2_3_03_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 3);
        }

        private bool CheckTask_2_3_04_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 3);
        }

        private bool CheckTask_2_3_05_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 3);
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

        static ExcelChecker2_3()
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
        // 共通プロセス・ヘルパーメソッド（ExcelChecker1_3を参考）
        // ==========================================
        
        private bool CheckTaskBasic(string filePath, string sheetName, Func<Worksheet, bool> checkLogic)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { return false; }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                worksheet = FindWorksheet(workbook, sheetName);
                if (worksheet == null) return false;

                return checkLogic(worksheet);
            }
            catch { return false; }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
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