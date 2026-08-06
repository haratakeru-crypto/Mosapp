using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace Libraries.Group2
{
    public class ExcelChecker2_10
    {
        public bool CheckTask_2_10_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_10_01_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_10_01_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_10_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_10_02_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_10_02_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_10_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_10_03_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_10_03_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_10_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_10_04_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_10_04_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_10_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_10_05_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_10_05_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_10_01_Impl(string filePath)
        {
            // CSVの解答手順: 「担当者マスター」シートの数式を表示
            // ExcelChecker1_9のCheckTask_1_9_01_Implを参考に、数式表示をチェック
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet originalSheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                // 元のアクティブシートを保存
                try { originalSheet = excelApp.ActiveSheet as Worksheet; }
                catch { }

                bool targetSheetCorrect = false;
                bool otherSheetsCorrect = true;

                // DisplayFormulasはWindowオブジェクトのプロパティ
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    try
                    {
                        ws.Activate();
                        bool isDisplayingFormulas = excelApp.ActiveWindow.DisplayFormulas;
                        
                        if (ws.Name == "担当者マスター")
                        {
                            if (isDisplayingFormulas) targetSheetCorrect = true;
                        }
                        else
                        {
                            if (isDisplayingFormulas) otherSheetsCorrect = false;
                        }
                    }
                    catch { }
                }

                // 元のシートに戻す
                if (originalSheet != null)
                {
                    try { originalSheet.Activate(); }
                    catch { }
                }

                return targetSheetCorrect && otherSheetsCorrect;
            }
            catch { return false; }
        }

        private bool CheckTask_2_10_02_Impl(string filePath)
        {
            // CSVの解答手順: 「売上一覧」シートの商品型番は降順、在庫は昇順に並べ替え
            // ExcelChecker1_9のCheckTask_1_9_02_Implを参考に、並べ替えをチェック
            return ProcessSheet(filePath, "売上一覧", (ws) =>
            {
                Range usedRange = ws.UsedRange;
                object[,] values = (object[,])usedRange.Value2;
                if (values == null) return false;

                int rowCount = values.GetLength(0);
                int colCount = values.GetLength(1);

                // ヘッダー行を探す
                int headerRow = -1;
                int colProductId = -1;
                int colStock = -1;

                for (int r = 1; r <= Math.Min(10, rowCount); r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        string val = Convert.ToString(values[r, c]);
                        if (val == "商品型番") colProductId = c;
                        if (val == "在庫") colStock = c;
                    }
                    if (colProductId != -1 && colStock != -1)
                    {
                        headerRow = r;
                        break;
                    }
                }

                if (headerRow == -1) return false;

                // データの並び順チェック
                for (int r = headerRow + 2; r <= rowCount; r++)
                {
                    string sIdCur = Convert.ToString(values[r, colProductId]);
                    string sIdPrev = Convert.ToString(values[r - 1, colProductId]);
                    
                    double dIdCur = 0, dIdPrev = 0;
                    bool isNum = double.TryParse(sIdCur, out dIdCur) && double.TryParse(sIdPrev, out dIdPrev);

                    int compareID;
                    if (isNum)
                    {
                        compareID = dIdCur.CompareTo(dIdPrev);
                    }
                    else
                    {
                        compareID = string.Compare(sIdCur, sIdPrev);
                    }

                    // 商品型番 (降順): 前の行 >= 今の行
                    if (compareID > 0)
                    {
                        Console.WriteLine($"[DEBUG] Task 2-10-2 Failed: Product ID not in descending order at row {r}.");
                        return false;
                    }

                    // 商品型番が同じ場合、在庫は昇順
                    if (compareID == 0 && colStock != -1)
                    {
                        string sStockCur = Convert.ToString(values[r, colStock]);
                        string sStockPrev = Convert.ToString(values[r - 1, colStock]);
                        
                        double dStockCur = 0, dStockPrev = 0;
                        bool isStockNum = double.TryParse(sStockCur, out dStockCur) && double.TryParse(sStockPrev, out dStockPrev);

                        int compareStock;
                        if (isStockNum)
                        {
                            compareStock = dStockCur.CompareTo(dStockPrev);
                        }
                        else
                        {
                            compareStock = string.Compare(sStockCur, sStockPrev);
                        }

                        // 在庫 (昇順): 前の行 <= 今の行
                        if (compareStock < 0)
                        {
                            Console.WriteLine($"[DEBUG] Task 2-10-2 Failed: Stock not in ascending order at row {r}.");
                            return false;
                        }
                    }
                }

                Console.WriteLine("[DEBUG] Task 2-10-2 Passed: Sort order is correct.");
                return true;
            });
        }

        private bool CheckTask_2_10_03_Impl(string filePath)
        {
            // CSVの解答手順: 「文化祭」シートの1日目から4日目に、条件付き書式で「3つの矢印（色分け）」を設定
            // ExcelChecker1_9のCheckTask_1_9_03_Implを参考に、ProcessSheetパターンを使用
            return ProcessSheet(filePath, "文化祭", (ws) =>
            {
                // 1日目から4日目の列の範囲を特定（例: D8:G16など）
                Range targetRange = ws.Range["D8:G16"];
                dynamic formatConditions = targetRange.FormatConditions;
                
                foreach (dynamic fc in formatConditions)
                {
                    try
                    {
                        int fcType = (int)fc.Type;
                        Console.WriteLine($"[DEBUG] Task 2-10-3: FormatCondition Type = {fcType}");
                        
                        // Type 6 = xlIconSet
                        if (fcType == 6)
                        {
                            dynamic iconSet = fc.IconSet;
                            int setId = (int)iconSet.ID;
                            
                            Console.WriteLine($"[DEBUG] Task 2-10-3: Found IconSet with ID = {setId}");

                            // ID 1 が「3つの矢印（色分け）」
                            if (setId == 1)
                            {
                                Console.WriteLine("[DEBUG] Task 2-10-3: Passed - IconSet ID = 1 found");
                                return true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[DEBUG] Task 2-10-3: Warning - {ex.Message}");
                    }
                }

                Console.WriteLine("[DEBUG] Task 2-10-3: Failed - No matching IconSet found");
                return false;
            });
        }

        private bool CheckTask_2_10_04_Impl(string filePath)
        {
            // CSVの解答手順: 「文化祭」シートの1日目から4日目に、条件付き書式で「35,000」より大きいセルに「濃い黄色の文字、黄色の背景」を設定
            // ExcelChecker1_9のCheckTask_1_9_04_Implを参考に、ProcessSheetパターンを使用
            return ProcessSheet(filePath, "文化祭", (ws) =>
            {
                dynamic usedRange = ws.UsedRange;
                dynamic formatConditions = usedRange.FormatConditions;

                foreach (dynamic fc in formatConditions)
                {
                    try
                    {
                        // Type 1 = xlCellValue, Operator 5 = xlGreater
                        if ((int)fc.Type == 1 && (int)fc.Operator == 5)
                        {
                            string f1 = "";
                            try { f1 = fc.Formula1; } catch {}
                            
                            // 35000より大きい条件をチェック
                            if (f1 == "=35000" || f1 == "35000" || f1 == "=35000" || f1.Contains("35000"))
                            {
                                Console.WriteLine("[DEBUG] Task 2-10-4 Passed.");
                                return true;
                            }
                        }
                    }
                    catch { }
                }
                Console.WriteLine("[DEBUG] Task 2-10-4 Failed: No matching conditional format found.");
                return false;
            });
        }

        private bool CheckTask_2_10_05_Impl(string filePath)
        {
            // CSVの解答手順: アクセシビリティチェックを行い、マイナスの通貨表示を「-6674」にする
            // ExcelChecker1_9のCheckTask_1_9_07_Implを参考に、NumberFormatLocalをチェック
            return ProcessSheet(filePath, "販売実績", (ws) =>
            {
                // F10セルをチェック（CSVの解答手順に記載されているセル）
                Range cell = ws.Range["F10"];
                
                // 表示形式を取得
                string numFormat = (string)cell.NumberFormatLocal;
                
                Console.WriteLine($"[DEBUG] Task 2-10-5: NumberFormat = '{numFormat}'");
                
                // Excelの表示形式の仕組み:
                // - "0" → 負の数は黒いマイナスで表示される → OK
                // - "[Red]0" → 負の数は赤色で表示される → NG
                // アクセシビリティチェックの要件: マイナスの通貨表示を「-6674」にする
                // つまり、赤色指定がなければ、負の数は黒いマイナス記号で表示される
                
                bool hasRedFormat = numFormat.Contains("[Red]") || numFormat.Contains("[赤]");
                
                if (hasRedFormat)
                {
                    Console.WriteLine("[DEBUG] Task 2-10-5: Failed - Red format found in NumberFormat");
                    return false;
                }
                
                // 赤色指定がなければOK
                Console.WriteLine("[DEBUG] Task 2-10-5: Passed - No red format, black minus will be displayed");
                return true;
            });
        }

        public bool CheckTask_2_10_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                bool existingCheck = CheckTask_2_10_06_Impl(filePath);
                bool comparisonCheck = CheckTask_2_10_06_Comparison(filePath);
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_10_06_Impl(string filePath)
        {
            // CSVの解答手順: 「売上一覧」シートのヘッダー右側にP/Nと表示、Pはページ番号、Nはページ数
            // ExcelChecker1_9のCheckTask_1_9_06_Implを参考に、フッターをチェック
            return ProcessSheet(filePath, "売上一覧", (ws) =>
            {
                string rightHeader = ws.PageSetup.RightHeader;
                // ヘッダー右側に&P/&Nが含まれているかチェック
                return !string.IsNullOrEmpty(rightHeader) && 
                       rightHeader.Contains("&P") && 
                       rightHeader.Contains("&N");
            });
        }

        public bool CheckTask_2_10_07()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                bool existingCheck = CheckTask_2_10_07_Impl(filePath);
                bool comparisonCheck = CheckTask_2_10_07_Comparison(filePath);
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_10_07_Impl(string filePath)
        {
            // CSVの解答手順: 「試験結果」シートのB5を基準にし、テキストファイル「結果内容」をインポート、受験番号、学籍番号が1行目のヘッダーになるように設定
            // ExcelChecker1_10のCheckTask_1_10_08_Implを参考に、テキストインポートをチェック
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

                // B5付近にQueryTableがあるかチェック
                if (worksheet.QueryTables.Count > 0)
                {
                    foreach (QueryTable qt in worksheet.QueryTables)
                    {
                        // クエリテーブルの接続文字列にテキストファイルが含まれているかチェック
                        string connection = qt.Connection as string;
                        if (connection != null && connection.Contains(".txt"))
                        {
                            // テーブルの開始位置がB5付近かチェック
                            Range destination = qt.Destination;
                            if (destination != null)
                            {
                                if (destination.Address.Contains("$B$5") || destination.Row == 5)
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }

                // ListObjectsもチェック（Power Queryでインポートした場合）
                if (worksheet.ListObjects.Count > 0)
                {
                    foreach (ListObject lo in worksheet.ListObjects)
                    {
                        Range headerRange = lo.HeaderRowRange;
                        if (headerRange != null && headerRange.Row >= 5 && headerRange.Column == 2)
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

        private bool CheckTask_2_10_06_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 10);
        }

        private bool CheckTask_2_10_07_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 10);
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
        private bool CheckTask_2_10_01_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 10);
        }

        private bool CheckTask_2_10_02_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 10);
        }

        private bool CheckTask_2_10_03_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 10);
        }

        private bool CheckTask_2_10_04_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 10);
        }

        private bool CheckTask_2_10_05_Comparison(string currentFilePath)
        {
            return PerformGeneralComparison(currentFilePath, 2, 10);
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

        static ExcelChecker2_10()
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