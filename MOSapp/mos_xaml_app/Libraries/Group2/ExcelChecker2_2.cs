using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

namespace Libraries.Group2
{
    public class ExcelChecker2_2
    {
        public bool CheckTask_2_2_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_2_01_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_2_01_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_2_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_2_02_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_2_02_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_2_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_2_03_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_2_03_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_2_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_2_04_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_2_04_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CheckTask_2_2_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                
                // 既存の細かい操作チェック
                bool existingCheck = CheckTask_2_2_05_Impl(filePath);
                
                // 比較ベースのチェック（CompletedとInitialを比較）
                bool comparisonCheck = CheckTask_2_2_05_Comparison(filePath);
                
                // 両方のチェックが成功した場合のみtrue
                return existingCheck && comparisonCheck;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_2_2_01_Impl(string filePath)
        {
            // ExcelChecker1_2のCheckTask_1_2_01_Implを参考に、ProcessTableTaskパターンを使用
            // タスク2-1: シート［試験結果］のテーブルの縞模様（行）を解除し、列の縞模様を設定
            return ProcessTableTask(filePath, "試験結果", (table) =>
            {
                // 1. プロパティチェック
                bool hasRowStripes = GetTableProperty<bool>(table, "ShowTableStyleRowStripes", true);
                bool hasColumnStripes = GetTableProperty<bool>(table, "ShowTableStyleColumnStripes", false);

                // 判定A: プロパティ設定が正しい (行OFF, 列ON)
                if (!hasRowStripes && hasColumnStripes)
                {
                    Console.WriteLine("[DEBUG] Task 2-1 Passed: Properties correct.");
                    return true;
                }

                // 2. 視覚的チェック（プロパティ取得失敗時の保険）
                bool visualRowBanding = CheckTableRowBandingVisually(table);
                bool visualColBanding = CheckTableColumnBandingVisually(table);
                
                // 判定B: 見た目が正しい (行の縞模様がなく、列の縞模様がある)
                bool isRowVisualOk = !visualRowBanding;
                bool isColVisualOk = visualColBanding;

                if (isRowVisualOk && isColVisualOk)
                {
                    Console.WriteLine("[DEBUG] Task 2-1 Passed: Visual check ok.");
                    return true;
                }

                Console.WriteLine("[DEBUG] Task 2-1 Failed.");
                return false;
            });
        }

        private bool CheckTask_2_2_02_Impl(string filePath)
        {
            // ExcelChecker1_2のCheckTask_1_2_02_Implを参考に、ProcessTableTaskパターンを使用
            // タスク2-2: シート［試験結果］のテーブルの最後の列を強調
            return ProcessTableTask(filePath, "試験結果", (table) =>
            {
                // 1. 「最初の列」がONになっていないかチェック (ONなら即不合格)
                bool isFirstColOn = GetTableProperty<bool>(table, "ShowTableStyleFirstColumn", false);
                if (isFirstColOn)
                {
                    Console.WriteLine("[DEBUG] Task 2-2 Failed: First Column emphasis is incorrectly ON.");
                    return false;
                }

                // 2. プロパティチェック: 「最後の列」がONか
                bool showLastColumn = GetTableProperty<bool>(table, "ShowTableStyleLastColumn", false);
                
                // 3. 視覚チェック: 最後の列が太字になっているか（色は2-1の縞模様で変わるため無視）
                bool visualLastEmphasis = CheckLastColumnEmphasisStrict(table);

                Console.WriteLine($"[DEBUG] ShowLastColProp: {showLastColumn}, VisualLastCol: {visualLastEmphasis}");

                // 判定: プロパティ設定または視覚的に正しいこと
                if (showLastColumn || visualLastEmphasis)
                {
                    Console.WriteLine("[DEBUG] Task 2-2 Passed.");
                    return true;
                }

                Console.WriteLine("[DEBUG] Task 2-2 Failed: Last column not emphasized.");
                return false;
            });
        }

        private bool CheckTask_2_2_03_Impl(string filePath)
        {
            // ExcelChecker1_2のCheckTask_1_2_03_Implを参考に、ProcessTableTaskパターンを使用
            // タスク2-3: シート［試験結果］のテーブルにテーブルスタイル「オレンジ、テーブルスタイル（中間）10」を設定
            return ProcessTableTask(filePath, "試験結果", (table) =>
            {
                string tableStyle = GetTableStyleName(table);
                Console.WriteLine($"[DEBUG] Current Style Name: {tableStyle}");

                // 厳密な判定リスト: 曖昧な "Medium", "10", "Orange" 単体を除外
                string[] strictValidStyles = {
                    "TableStyleMedium10",        // 内部名(英語)
                    "Medium10",                  // 短縮名
                    "Medium 10",                 // 空白あり
                    "テーブルスタイル（中間）10", // 日本語正式名
                    "10"                         // 末尾一致用
                };

                bool matchFound = false;
                foreach (string validStyle in strictValidStyles)
                {
                    // 完全一致または、末尾が明確に一致することを確認
                    if (!string.IsNullOrEmpty(tableStyle) && 
                        tableStyle.EndsWith(validStyle, StringComparison.OrdinalIgnoreCase))
                    {
                        matchFound = true;
                        break;
                    }
                }

                if (matchFound)
                {
                    Console.WriteLine("[DEBUG] Task 2-3 Passed: Style matches.");
                    return true;
                }

                Console.WriteLine("[DEBUG] Task 2-3 Failed: Style name does not match required 'Medium 10'.");
                return false;
            });
        }

        private bool CheckTask_2_2_04_Impl(string filePath)
        {
            // ExcelChecker1_2のCheckTask_1_2_04_Implを参考に、ProcessTableTaskパターンを使用
            // タスク2-4: シート「担当者マスター」のテーブルを有楽町店の担当だけ表示（模試②用：解答手順あり模擬試験①問題文.csvを参照）
            return ProcessTableTask(filePath, "担当者マスター", (table) =>
            {
                // 担当者列または店舗列を探す（「有楽町店」を含む列を探す）
                int storeColIndex = -1;
                for (int i = 1; i <= table.ListColumns.Count; i++)
                {
                    string colName = table.ListColumns[i].Name;
                    // 「店」を含む列を探す（例：「店舗」「担当店舗」など）
                    if (colName.Contains("店") || colName.Contains("担当"))
                    {
                        storeColIndex = i;
                        break;
                    }
                }

                // 店舗列が見つからない場合、データ行から「有楽町店」を含む列を探す
                if (storeColIndex == -1)
                {
                    Range tempDataBody = table.DataBodyRange;
                    if (tempDataBody != null && tempDataBody.Rows.Count > 0)
                    {
                        Range firstRow = tempDataBody.Rows[1];
                        for (int i = 1; i <= table.ListColumns.Count; i++)
                        {
                            string cellValue = ((Range)firstRow.Cells[1, i]).Value2?.ToString() ?? "";
                            if (cellValue.Contains("有楽町店"))
                            {
                                storeColIndex = i;
                                break;
                            }
                        }
                    }
                }

                if (storeColIndex == -1)
                {
                    Console.WriteLine("[DEBUG] Task 2-4 Failed: Column containing '有楽町店' not found.");
                    return false;
                }

                // AutoFilterがOFFなら即不合格
                if (table.AutoFilter == null)
                {
                    Console.WriteLine("[DEBUG] Task 2-4 Failed: AutoFilter is not enabled.");
                    return false;
                }

                // 実際のフィルタリング状態を確認
                Range dataBody = table.DataBodyRange;
                if (dataBody == null) return false;

                bool wrongDataFound = false;
                bool correctDataFound = false;
                int visibleRows = 0;

                foreach (Range row in dataBody.Rows)
                {
                    if (!(bool)row.EntireRow.Hidden)
                    {
                        visibleRows++;
                        string val = ((Range)row.Cells[1, storeColIndex]).Value2?.ToString() ?? "";
                        if (!val.Contains("有楽町店"))
                        {
                            wrongDataFound = true; // 有楽町店以外が見えている＝不正解
                            Console.WriteLine($"[DEBUG] Found wrong visible data: {val}");
                            break;
                        }
                        else
                        {
                            correctDataFound = true;
                        }
                    }
                }

                // 厳密な判定:
                // - 間違ったデータが見えていないこと
                // - 正しいデータが少なくとも1つ見えていること
                if (!wrongDataFound && correctDataFound)
                {
                    Console.WriteLine("[DEBUG] Task 2-4 Passed: Only '有楽町店' is visible.");
                    return true;
                }

                Console.WriteLine("[DEBUG] Task 2-4 Failed.");
                return false;
            }, sheetNameOrNull: "担当者マスター");
        }

        private bool CheckTask_2_2_05_Impl(string filePath)
        {
            // ExcelChecker1_2のCheckTask_1_2_05_Implを参考に、ProcessTableTaskパターンを使用
            // タスク2-5: シート「文化祭」のテーブルに4日間合計の列を追加（模試②用：解答手順あり模擬試験①問題文.csvを参照）
            return ProcessTableTask(filePath, "文化祭", (table) =>
            {
                // 「4日間合計」または「合計」という名前の列が存在するかチェック
                bool hasTotalColumn = false;
                string totalColumnName = "";
                
                for (int i = 1; i <= table.ListColumns.Count; i++)
                {
                    string colName = table.ListColumns[i].Name;
                    if (colName.Contains("4日間合計") || colName.Contains("合計") || colName.Contains("4日"))
                    {
                        hasTotalColumn = true;
                        totalColumnName = colName;
                        break;
                    }
                }

                if (!hasTotalColumn)
                {
                    Console.WriteLine("[DEBUG] Task 2-5 Failed: Column '4日間合計' not found.");
                    return false;
                }

                Console.WriteLine($"[DEBUG] Task 2-5 Passed: Found total column '{totalColumnName}'.");
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
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク01）
        /// </summary>
        private bool CheckTask_2_2_01_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;
            Range currentCell = null;
            Range completedCell = null;
            Comment currentComment = null;
            Comment completedComment = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 2);
                string completedFilePath = GetCompletedDataFilePath(2, 2);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "売上一覧");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "売上一覧");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                // G4セルのメモを比較
                currentCell = currentWorksheet.Range["G4"];
                completedCell = completedWorksheet.Range["G4"];

                if (currentCell == null || completedCell == null)
                    return false;

                currentComment = currentCell.Comment;
                completedComment = completedCell.Comment;

                if ((currentComment == null) != (completedComment == null))
                    return false;

                if (currentComment == null && completedComment == null)
                    return true;

                string currentText = currentComment.Text() ?? string.Empty;
                string completedText = completedComment.Text() ?? string.Empty;

                string normalizedCurrent = currentText.Trim().Replace("\r", "").Replace("\n", "");
                string normalizedCompleted = completedText.Trim().Replace("\r", "").Replace("\n", "");

                return normalizedCurrent == normalizedCompleted;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComparisonComObject(completedComment);
                ReleaseComparisonComObject(currentComment);
                ReleaseComparisonComObject(completedCell);
                ReleaseComparisonComObject(currentCell);
                ReleaseComparisonComObject(completedWorksheet);
                ReleaseComparisonComObject(currentWorksheet);
                CloseComparisonWorkbook(completedWorkbook);
            }
        }

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク02）
        /// </summary>
        private bool CheckTask_2_2_02_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 2);
                string completedFilePath = GetCompletedDataFilePath(2, 2);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "売上一覧");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "売上一覧");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                return currentWorksheet.PageSetup.Orientation == completedWorksheet.PageSetup.Orientation;
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

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク03）
        /// </summary>
        private bool CheckTask_2_2_03_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 2);
                string completedFilePath = GetCompletedDataFilePath(2, 2);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "売上一覧");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "売上一覧");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                string currentPrintArea = currentWorksheet.PageSetup.PrintArea ?? string.Empty;
                string completedPrintArea = completedWorksheet.PageSetup.PrintArea ?? string.Empty;

                string normalizedCurrent = NormalizeRange(currentPrintArea);
                string normalizedCompleted = NormalizeRange(completedPrintArea);

                return normalizedCurrent == normalizedCompleted;
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

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク04）
        /// </summary>
        private bool CheckTask_2_2_04_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 2);
                string completedFilePath = GetCompletedDataFilePath(2, 2);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "売上一覧");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "売上一覧");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

                string currentTitleRows = currentWorksheet.PageSetup.PrintTitleRows ?? string.Empty;
                string completedTitleRows = completedWorksheet.PageSetup.PrintTitleRows ?? string.Empty;

                string normalizedCurrent = NormalizeRange(currentTitleRows);
                string normalizedCompleted = NormalizeRange(completedTitleRows);

                return normalizedCurrent == normalizedCompleted;
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

        /// <summary>
        /// 比較ベースのチェック：CompletedとInitialフォルダを比較（タスク05）
        /// </summary>
        private bool CheckTask_2_2_05_Comparison(string currentFilePath)
        {
            Application excelApp = null;
            Workbook currentWorkbook = null;
            Workbook completedWorkbook = null;
            Worksheet currentWorksheet = null;
            Worksheet completedWorksheet = null;

            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                string initialFilePath = GetInitialDataFilePath(2, 2);
                string completedFilePath = GetCompletedDataFilePath(2, 2);

                if (string.IsNullOrEmpty(initialFilePath) || string.IsNullOrEmpty(completedFilePath))
                {
                    return true;
                }

                currentWorkbook = GetComparisonWorkbook(excelApp, currentFilePath);
                completedWorkbook = GetComparisonWorkbook(excelApp, completedFilePath);

                if (currentWorkbook == null || completedWorkbook == null)
                    return false;

                currentWorksheet = FindComparisonWorksheet(currentWorkbook, "販売実績");
                completedWorksheet = FindComparisonWorksheet(completedWorkbook, "販売実績");

                if (currentWorksheet == null || completedWorksheet == null)
                    return false;

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

        static ExcelChecker2_2()
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

        private static Worksheet FindComparisonWorksheet(Workbook workbook, string sheetName)
        {
            if (workbook == null) return null;
            try
            {
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return worksheet;
                    }
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
        // 共通プロセス・ヘルパーメソッド（ExcelChecker1_2を参考）
        // ==========================================
        
        // リソース管理とテーブル取得を共通化するラッパー
        private bool ProcessTableTask(string filePath, string sheetName, Func<ListObject, bool> checkLogic, string sheetNameOrNull = null)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            
            string targetSheet = sheetNameOrNull ?? sheetName;

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
                    if (ws.Name.Equals(targetSheet, StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = ws;
                        break;
                    }
                }
                if (worksheet == null) return false;

                ListObject table = null;
                
                if (targetSheet == "担当者マスター")
                {
                    // 担当者マスターシートのテーブルを探す（「有楽町店」を含む列があるテーブル）
                    foreach (ListObject t in worksheet.ListObjects)
                    {
                        try 
                        { 
                            // 「店」または「担当」を含む列があるテーブルを探す
                            bool found = false;
                            for (int i = 1; i <= t.ListColumns.Count; i++)
                            {
                                string colName = t.ListColumns[i].Name;
                                if (colName.Contains("店") || colName.Contains("担当"))
                                {
                                    found = true;
                                    break;
                                }
                            }
                            if (found)
                            {
                                table = t; 
                                break; 
                            }
                        } 
                        catch {}
                    }
                    // 見つからない場合は最初のテーブルを使用
                    if (table == null && worksheet.ListObjects.Count > 0)
                    {
                        table = worksheet.ListObjects[1];
                    }
                }
                else
                {
                    if (worksheet.ListObjects.Count > 0) table = worksheet.ListObjects[1];
                }

                if (table == null) return false;

                return checkLogic(table);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error in ProcessTableTask: {ex.Message}");
                return false;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        // テーブルスタイル名を取得
        private string GetTableStyleName(ListObject table)
        {
            try
            {
                dynamic style = table.TableStyle;
                return style.Name;
            }
            catch { return ""; }
        }

        // COMオブジェクトのプロパティを安全に取得
        private T GetTableProperty<T>(ListObject table, string propertyName, T defaultValue = default(T))
        {
            try
            {
                // リフレクションでプロパティ値取得
                var prop = table.GetType().InvokeMember(propertyName, System.Reflection.BindingFlags.GetProperty, null, table, null);
                return (T)Convert.ChangeType(prop, typeof(T));
            }
            catch
            {
                return defaultValue;
            }
        }

        // 行の縞模様の視覚的チェック（DisplayFormat.Interior.Colorを使用）
        private bool CheckTableRowBandingVisually(ListObject table)
        {
            try
            {
                Range data = table.DataBodyRange;
                if (data == null || data.Rows.Count < 3) return false;
                
                var colors = new List<double>();
                for(int i=1; i<=Math.Min(data.Rows.Count, 10); i++)
                {
                    Range cell = (Range)data.Cells[i, 1];
                    colors.Add((double)cell.DisplayFormat.Interior.Color);
                }

                int changes = 0;
                for(int i=1; i<colors.Count; i++) if(colors[i] != colors[i-1]) changes++;

                return changes >= 3;
            }
            catch { return false; }
        }

        // 列の縞模様の視覚的チェック（DisplayFormat.Interior.Colorを使用）
        private bool CheckTableColumnBandingVisually(ListObject table)
        {
            try
            {
                Range data = table.DataBodyRange;
                if (data == null || data.Columns.Count < 3) return false;

                var colors = new List<double>();
                for (int i = 1; i <= Math.Min(data.Columns.Count, 10); i++)
                {
                    Range cell = (Range)data.Cells[1, i];
                    colors.Add((double)cell.DisplayFormat.Interior.Color);
                }

                int changes = 0;
                for (int i = 1; i < colors.Count; i++) if (colors[i] != colors[i - 1]) changes++;

                return changes >= 2; // 列は数が少ないので閾値を下げる
            }
            catch { return false; }
        }

        // 最後の列の強調チェック
        private bool CheckLastColumnEmphasisStrict(ListObject table)
        {
            try
            {
                Range data = table.DataBodyRange;
                if (data == null || data.Columns.Count < 1) return false;
                
                int lastCol = data.Columns.Count;
                // 1行目の最後の列のセルを取得
                Range lastCell = (Range)data.Cells[1, lastCol];
                
                // 色は2-1の列の縞模様で変わるため判定には使わない。
                // 「最後の列」オプションは太字になるため、太字かどうかで判定する。
                dynamic lastFont = lastCell.Font;
                bool isBold = lastFont.Bold;

                return isBold;
            }
            catch { return false; }
        }
    }
}
