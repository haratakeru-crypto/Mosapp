using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_2
    {
        public bool CheckExcel(string filePath)
        {
            return true;
        }

        // ラッパーメソッド群
        public bool CheckTask_1_2_01() => RunCheck(CheckTask_1_2_01_Impl, "Task 1 (Stripes)");
        public bool CheckTask_1_2_02() => RunCheck(CheckTask_1_2_02_Impl, "Task 2 (Last Col)");
        public bool CheckTask_1_2_03() => RunCheck(CheckTask_1_2_03_Impl, "Task 3 (Style)");
        public bool CheckTask_1_2_04() => RunCheck(CheckTask_1_2_04_Impl, "Task 4 (Filter)");
        public bool CheckTask_1_2_05() => RunCheck(CheckTask_1_2_05_Impl, "Task 5 (Resize)");

        // config tabs["1"] project 1 用エイリアス
        public bool CheckTask_1_1_01() => CheckTask_1_2_01();
        public bool CheckTask_1_1_02() => CheckTask_1_2_02();
        public bool CheckTask_1_1_03() => CheckTask_1_2_03();
        public bool CheckTask_1_1_04() => CheckTask_1_2_04();
        public bool CheckTask_1_1_05() => CheckTask_1_2_05();

        // 共通エラーハンドリング用ヘルパー
        private bool RunCheck(Func<string, bool> checkImpl, string taskName)
        {
            try
            {
                Console.WriteLine($"[DEBUG] {taskName} called");
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return checkImpl(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Exception in {taskName}: {ex.Message}");
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        // ==========================================
        // タスク1: 縞模様（行）を解除し、縞模様（列）を設定
        // ==========================================
        private bool CheckTask_1_2_01_Impl(string filePath)
        {
            return ProcessTableTask(filePath, "試験結果", (table) =>
            {
                bool hasRowStripes = GetTableProperty<bool>(table, "ShowTableStyleRowStripes", true);
                bool hasColumnStripes = GetTableProperty<bool>(table, "ShowTableStyleColumnStripes", false);

                if (!hasRowStripes && hasColumnStripes)
                {
                    Console.WriteLine("[DEBUG] Task 1 Passed: Properties correct.");
                    return true;
                }

                bool visualRowBanding = CheckTableRowBandingVisually(table);
                bool visualColBanding = CheckTableColumnBandingVisually(table);
                bool isRowVisualOk = !visualRowBanding;
                bool isColVisualOk = visualColBanding;

                if (isRowVisualOk && isColVisualOk)
                {
                    Console.WriteLine("[DEBUG] Task 1 Passed: Visual check ok.");
                    return true;
                }

                bool rowStillOn = hasRowStripes || visualRowBanding;
                bool colStillOff = !hasColumnStripes && !visualColBanding;
                if (rowStillOn)
                    ExcelScoreExplanation.Note("テーブルの縞模様（行）が解除されていません。");
                if (colStillOff)
                    ExcelScoreExplanation.Note("テーブルの縞模様（列）が設定されていません。");
                if (!rowStillOn && !colStillOff)
                    ExcelScoreExplanation.Note("テーブルの縞模様（行を解除、列を設定）になっていません。");
                return false;
            });
        }

        // ==========================================
        // タスク2: 最後の列を強調 (修正版)
        // ==========================================
        private bool CheckTask_1_2_02_Impl(string filePath)
        {
            return ProcessTableTask(filePath, "試験結果", (table) =>
            {
                bool isFirstColOn = GetTableProperty<bool>(table, "ShowTableStyleFirstColumn", false);
                if (isFirstColOn)
                    ExcelScoreExplanation.Note("最初の列が強調されています。");

                bool showLastColumn = GetTableProperty<bool>(table, "ShowTableStyleLastColumn", false);
                if (showLastColumn)
                {
                    if (isFirstColOn)
                        return false;
                    Console.WriteLine("[DEBUG] Task 2 Passed: ShowTableStyleLastColumn is ON.");
                    return true;
                }

                bool lastColOk = false;
                if (TryGetLastColumnVisualEmphasis(table, out bool colorDiff, out bool isBold))
                {
                    bool hasColumnStripes = GetTableProperty<bool>(table, "ShowTableStyleColumnStripes", false);
                    // 列の縞模様による色差だけでは「最後の列の強調」とみなさない
                    lastColOk = !(hasColumnStripes && colorDiff && !isBold);
                    if (lastColOk)
                        Console.WriteLine($"[DEBUG] Task 2: VisualLastCol colorDiff={colorDiff}, bold={isBold}");
                }

                if (!lastColOk)
                    ExcelScoreExplanation.Note("最後の列が強調されていません。");

                if (!isFirstColOn && lastColOk)
                {
                    Console.WriteLine("[DEBUG] Task 2 Passed.");
                    return true;
                }

                return false;
            });
        }

        // ==========================================
        // タスク3: スタイル「オレンジ、テーブルスタイル（中間）10」
        // ==========================================
        private bool CheckTask_1_2_03_Impl(string filePath)
        {
            return ProcessTableTask(filePath, "試験結果", (table) =>
            {
                string tableStyle = GetTableStyleName(table);
                Console.WriteLine($"[DEBUG] Current Style Name: {tableStyle}");

                string[] strictValidStyles = {
                    "TableStyleMedium10",
                    "Medium10",
                    "Medium 10",
                    "テーブルスタイル（中間）10",
                    "TableStyleMedium10"
                };

                foreach (string validStyle in strictValidStyles)
                {
                    if (!string.IsNullOrEmpty(tableStyle) &&
                        tableStyle.EndsWith(validStyle, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine("[DEBUG] Task 3 Passed: Style matches.");
                        return true;
                    }
                }

                if (string.IsNullOrEmpty(tableStyle))
                    return Miss("テーブルスタイルが設定されていません。");
                return Miss($"テーブルスタイルが「{Quote(DescribeTableStyle(tableStyle))}」になっています。");
            });
        }

        // ==========================================
        // タスク4: フィルター「法学科」
        // ==========================================
        private bool CheckTask_1_2_04_Impl(string filePath)
        {
            return ProcessTableTask(filePath, "担当者リスト", (table) =>
            {
                int gakkaColIndex = -1;
                for (int i = 1; i <= table.ListColumns.Count; i++)
                {
                    if (table.ListColumns[i].Name.Contains("学科"))
                    {
                        gakkaColIndex = i;
                        break;
                    }
                }

                if (gakkaColIndex == -1)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                if (table.AutoFilter == null)
                    return Miss("フィルターが使われていません。");

                Range dataBody = table.DataBodyRange;
                if (dataBody == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                bool correctDataFound = false;
                int visibleRows = 0;
                var wrongValues = new List<string>();

                foreach (Range row in dataBody.Rows)
                {
                    if ((bool)row.EntireRow.Hidden)
                        continue;

                    visibleRows++;
                    string val = ((Range)row.Cells[1, gakkaColIndex]).Value2?.ToString() ?? "";
                    if (!val.Contains("法学科"))
                    {
                        string shown = Quote(string.IsNullOrWhiteSpace(val) ? "（空）" : val.Trim());
                        if (!wrongValues.Contains(shown))
                            wrongValues.Add(shown);
                    }
                    else
                    {
                        correctDataFound = true;
                    }
                }

                if (wrongValues.Count == 0 && correctDataFound)
                {
                    Console.WriteLine("[DEBUG] Task 4 Passed: Only '法学科' is visible.");
                    return true;
                }

                if (visibleRows == 0)
                {
                    ExcelScoreExplanation.Note("フィルター後に表示されている行がありません。");
                    return false;
                }

                if (wrongValues.Count > 0)
                    ExcelScoreExplanation.Note($"法学科以外の行（学科が「{JoinNames(wrongValues)}」）も表示されています。");
                if (!correctDataFound)
                    ExcelScoreExplanation.Note("法学科の行が表示されていません。");
                return false;
            }, sheetNameOrNull: "担当者リスト");
        }

        // ==========================================
        // タスク5: 合計列追加・範囲変更 (修正版・残留チェック強化)
        // ==========================================
        private bool CheckTask_1_2_05_Impl(string filePath)
        {
            return ProcessTableTask(filePath, "イベント売上", (table) =>
            {
                string rangeAddr = table.Range.Address.Replace("$", "").Replace(" ", "").ToUpper();
                bool rangeOk = rangeAddr.EndsWith("G16") && table.Range.Columns.Count == 7;
                if (!rangeOk)
                    ExcelScoreExplanation.Note($"テーブルの範囲が「{Quote(rangeAddr)}」になっています。");

                bool residualOk = true;
                try
                {
                    Worksheet ws = table.Parent as Worksheet;
                    Range ghostRange = ws.Range["H4:H16"];
                    var residual = new List<string>();

                    foreach (Range cell in ghostRange)
                    {
                        string address = cell.Address[false, false];
                        if ((int)cell.Interior.ColorIndex != -4142)
                        {
                            if (!residual.Contains(address))
                                residual.Add(address);
                            continue;
                        }

                        int[] bordersToCheck = { 10, 8, 9 };
                        foreach (int borderIndex in bordersToCheck)
                        {
                            Border border = cell.Borders[(XlBordersIndex)borderIndex];
                            if ((int)border.LineStyle != -4142)
                            {
                                if (!residual.Contains(address))
                                    residual.Add(address);
                                break;
                            }
                        }
                    }

                    if (residual.Count > 0)
                    {
                        residualOk = false;
                        ExcelScoreExplanation.Note($"テーブル外の{JoinNames(residual)}に書式が残っています。");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[DEBUG] Error checking ghost range: {ex.Message}");
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }

                if (rangeOk && residualOk)
                {
                    Console.WriteLine("[DEBUG] Task 5 Passed: Range correct and clean.");
                    return true;
                }

                return false;
            });
        }

        private string GetCurrentExcelFilePath()
        {
            try
            {
                Application excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                // 固定ファイル名に依存せず、現在アクティブなブックを採点対象にする。
                return excelApp.ActiveWorkbook?.FullName;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error: {ex.Message}");
                return null;
            }
        }

        // ==========================================
        // 共通プロセス・ヘルパーメソッド
        // ==========================================

        // リソース管理とテーブル取得を共通化するラッパー
        private bool ProcessTableTask(string filePath, string sheetName, Func<ListObject, bool> checkLogic, string sheetNameOrNull = null)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            
            // Task4のように引数で指定されたシート名を使うか、デフォルトを使うか
            string targetSheet = sheetNameOrNull ?? sheetName;

            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { excelApp = new Application { Visible = false }; }

                // ワークブック特定
                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                // ワークシート特定 (Task4の場合、部分一致などで探すロジックが必要ならここに実装)
                foreach(Worksheet ws in workbook.Worksheets) {
                    if(ws.Name == targetSheet) { worksheet = ws; break; }
                }
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                // テーブル特定 (Task4の「学科」列を持つテーブル検索ロジックはTask4内に記述推奨だが、ここでは簡易的に1つ目を取得)
                // 注: 元コードに合わせて ListObjects[1] を基本としますが、Task4用にロジック分岐が必要
                ListObject table = null;
                
                if (targetSheet == "担当者リスト") // Task4用
                {
                    foreach (ListObject t in worksheet.ListObjects) {
                        // 学科列があるかチェック
                        try { if(t.ListColumns["学科"] != null) table = t; } catch {}
                        if(table != null) break;
                    }
                }
                else
                {
                    if (worksheet.ListObjects.Count > 0) table = worksheet.ListObjects[1];
                }

                if (table == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                // 実際の判定ロジックを実行
                return checkLogic(table);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error in ProcessTableTask: {ex.Message}");
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                // Workbook, Appは閉じない（テスト継続のため）
            }
        }

        private Worksheet FindWorksheet(Workbook workbook, string worksheetName)
        {
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return ws;
                }
            }
            return null;
        }

        private ListObject FindTable(Worksheet worksheet)
        {
            if (worksheet.ListObjects.Count > 0)
            {
                return worksheet.ListObjects[1];
            }
            return null;
        }



        // --- 以下、既存のヘルパーメソッドの厳格版 ---

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
                    
                    // 修正後: DisplayFormat.Interior.Color (見た目通りの色を取得する)
                    colors.Add((double)cell.DisplayFormat.Interior.Color);
                }

                // 色の変化回数をカウント
                int changes = 0;
                for(int i=1; i<colors.Count; i++) if(colors[i] != colors[i-1]) changes++;

                // 縞模様なら、行数に応じて頻繁に色が切り替わるはず
                // 10行チェックして変化が3回以上あれば縞模様とみなす（タイトル行除く）
                return changes >= 3;
            }
            catch { return false; }
        }

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
                    
                    // 修正後: DisplayFormat.Interior.Color (見た目通りの色を取得する)
                    colors.Add((double)cell.DisplayFormat.Interior.Color);
                }

                int changes = 0;
                for (int i = 1; i < colors.Count; i++) if (colors[i] != colors[i - 1]) changes++;

                return changes >= 2; // 列は数が少ないので閾値を下げる
            }
            catch { return false; }
        }

        private string GetTableStyleName(ListObject table)
        {
            try
            {
                dynamic style = table.TableStyle;
                return style.Name;
            }
            catch { return ""; }
        }

        private T GetTableProperty<T>(ListObject table, string propertyName, T defaultValue)
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


        private bool CheckLastColumnEmphasis(ListObject table)
        {
            try
            {
                Range data = table.DataBodyRange;
                int lastCol = data.Columns.Count;
                int checkRow = 1; 
                
                Range firstCell = (Range)data.Cells[checkRow, 1];
                Range lastCell = (Range)data.Cells[checkRow, lastCol];
                
                // 修正後: DisplayFormat.Interior.Color (見た目通りの色を取得する)
                double firstColColor = (double)firstCell.DisplayFormat.Interior.Color;
                double lastColColor = (double)lastCell.DisplayFormat.Interior.Color;
                
                // 太字チェックも追加
                dynamic lastFont = lastCell.Font;
                bool isBold = lastFont.Bold;

                // 色が違う OR 太字になっている
                return (firstColColor != lastColColor) || isBold;
            }
            catch { return false; }
        }

        private bool CheckLastColumnEmphasisStrict(ListObject table)
        {
            return TryGetLastColumnVisualEmphasis(table, out bool colorDiff, out bool isBold)
                && (colorDiff || isBold);
        }

        private bool TryGetLastColumnVisualEmphasis(ListObject table, out bool colorDiff, out bool isBold)
        {
            colorDiff = false;
            isBold = false;
            try
            {
                Range data = table.DataBodyRange;
                if (data == null || data.Columns.Count < 2) return false;

                int lastCol = data.Columns.Count;
                int checkRow = 1;
                int normalCol = Math.Min(2, lastCol - 1);

                Range normalCell = (Range)data.Cells[checkRow, normalCol];
                Range lastCell = (Range)data.Cells[checkRow, lastCol];

                double normalColColor = (double)normalCell.DisplayFormat.Interior.Color;
                double lastColColor = (double)lastCell.DisplayFormat.Interior.Color;
                colorDiff = normalColColor != lastColColor;

                dynamic lastFont = lastCell.Font;
                isBold = lastFont.Bold;

                return colorDiff || isBold;
            }
            catch { return false; }
        }

        private static bool Miss(string reason)
        {
            ExcelScoreExplanation.Note(reason);
            return false;
        }

        private static string Quote(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "（空）";
            string text = value.Replace("\r", "").Replace("\n", " ");
            const int maxLen = 40;
            if (text.Length <= maxLen)
                return text;
            return text.Substring(0, maxLen) + "…";
        }

        private static string JoinNames(IList<string> names)
        {
            if (names == null || names.Count == 0)
                return "";
            const int maxItems = 5;
            if (names.Count <= maxItems)
                return string.Join("、", names);
            return string.Join("、", names.Take(maxItems)) + "ほか";
        }

        private static string DescribeTableStyle(string styleName)
        {
            if (string.IsNullOrEmpty(styleName))
                return "（なし）";
            if (styleName.IndexOf("TableStyleMedium", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                string num = styleName.Replace("TableStyleMedium", "").Replace("tableStyleMedium", "");
                return "テーブルスタイル（中間）" + num;
            }
            if (styleName.IndexOf("TableStyleLight", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                string num = styleName.Replace("TableStyleLight", "").Replace("tableStyleLight", "");
                return "テーブルスタイル（淡色）" + num;
            }
            if (styleName.IndexOf("TableStyleDark", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                string num = styleName.Replace("TableStyleDark", "").Replace("tableStyleDark", "");
                return "テーブルスタイル（濃色）" + num;
            }
            return styleName;
        }

    }
}