using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace Libraries.Group1
{
    public class ExcelChecker1_2_PV1
    {
        public bool CheckExcel(string filePath)
        {
            return true;
        }

        // ラッパーメソッド群
        public bool CheckTask_1_2_01() => RunCheck(CheckTask_1_2_01_Impl, "PV1 Task 1 (Row Stripes)");
        public bool CheckTask_1_2_02() => RunCheck(CheckTask_1_2_02_Impl, "PV1 Task 2 (First Col)");
        public bool CheckTask_1_2_03() => RunCheck(CheckTask_1_2_03_Impl, "PV1 Task 3 (Style M11)");
        public bool CheckTask_1_2_04() => RunCheck(CheckTask_1_2_04_Impl, "PV1 Task 4 (Filter 経済)");
        public bool CheckTask_1_2_05() => RunCheck(CheckTask_1_2_05_Impl, "PV1 Task 5 (Resize)");

        // 共通エラーハンドリング用ヘルパー
        private bool RunCheck(Func<string, bool> checkImpl, string taskName)
        {
            try
            {
                Console.WriteLine($"[DEBUG] {taskName} called");
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath)) return false;
                return checkImpl(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Exception in {taskName}: {ex.Message}");
                return false;
            }
        }

        // ==========================================
        // タスク1: 縞模様（列）を解除し、縞模様（行）を設定（類題1）
        // ==========================================
        private bool CheckTask_1_2_01_Impl(string filePath)
        {
            return ProcessTableTask(filePath, "試験結果", (table) =>
            {
                bool hasRowStripes = GetTableProperty<bool>(table, "ShowTableStyleRowStripes", false);
                bool hasColumnStripes = GetTableProperty<bool>(table, "ShowTableStyleColumnStripes", true);

                if (hasRowStripes && !hasColumnStripes)
                {
                    Console.WriteLine("[DEBUG] PV1 Task 1 Passed: Properties correct.");
                    return true;
                }

                bool visualRowBanding = CheckTableRowBandingVisually(table);
                bool visualColBanding = CheckTableColumnBandingVisually(table);

                if (visualRowBanding && !visualColBanding)
                {
                    Console.WriteLine("[DEBUG] PV1 Task 1 Passed: Visual check ok.");
                    return true;
                }

                return false;
            });
        }

        // ==========================================
        // タスク2: 最初の列を強調（類題1）
        // ==========================================
        private bool CheckTask_1_2_02_Impl(string filePath)
        {
            return ProcessTableTask(filePath, "試験結果", (table) =>
            {
                bool isLastColOn = GetTableProperty<bool>(table, "ShowTableStyleLastColumn", false);
                if (isLastColOn)
                {
                    Console.WriteLine("[DEBUG] PV1 Task 2 Failed: Last Column emphasis is incorrectly ON.");
                    return false;
                }

                bool showFirstColumn = GetTableProperty<bool>(table, "ShowTableStyleFirstColumn", false);
                bool visualFirstEmphasis = CheckFirstColumnEmphasisStrict(table);

                Console.WriteLine($"[DEBUG] ShowFirstColProp: {showFirstColumn}, VisualFirstCol: {visualFirstEmphasis}");

                if (showFirstColumn || visualFirstEmphasis)
                {
                    Console.WriteLine("[DEBUG] PV1 Task 2 Passed.");
                    return true;
                }

                Console.WriteLine("[DEBUG] PV1 Task 2 Failed: First column not emphasized.");
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

                // 厳密な判定リスト: 曖昧な "Medium", "10", "Orange" 単体を除外
                string[] strictValidStyles = {
                    "TableStyleMedium11",
                    "Medium11",
                    "Medium 11",
                    "テーブルスタイル（中間）11"
                };

                bool matchFound = false;
                foreach (string validStyle in strictValidStyles)
                {
                    // 完全一致または、末尾が明確に一致することを確認
                    // "BlueMedium10" などが "Medium10" にヒットしないように注意が必要だが、
                    // Excelの内部名体系的に末尾一致でほぼ特定可能
                    if (!string.IsNullOrEmpty(tableStyle) && 
                        tableStyle.EndsWith(validStyle, StringComparison.OrdinalIgnoreCase))
                    {
                        matchFound = true;
                        break;
                    }
                }

                // 色による補完チェック（スタイル名が取得できない場合のみ）
                if (!matchFound)
                {
                    // オレンジ色(RGB)が含まれているか厳密にチェックするロジックがあれば良いが、
                    // ここではスタイル名の不一致は不合格とする（厳格化のため）
                    Console.WriteLine("[DEBUG] Task 3 Failed: Style name does not match required 'Medium 11'.");
                    return false;
                }

                Console.WriteLine("[DEBUG] Task 3 Passed: Style matches.");
                return true;
            });
        }

        // ==========================================
        // タスク4: フィルター「経済学科」
        // ==========================================
        private bool CheckTask_1_2_04_Impl(string filePath)
        {
            // ワークシートとテーブル名が異なるため個別実装
            return ProcessTableTask(filePath, "担当者リスト", (table) =>
            {
                // 学科列を探す
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
                {
                    Console.WriteLine("[DEBUG] Task 4 Failed: Column '学科' not found.");
                    return false;
                }

                // AutoFilterがOFFなら即不合格
                if (table.AutoFilter == null)
                {
                    Console.WriteLine("[DEBUG] Task 4 Failed: AutoFilter is not enabled.");
                    return false;
                }

                // 実際のフィルタリング状態を確認

                // 2. 可視行の実データ確認 (これが最重要)
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
                        string val = ((Range)row.Cells[1, gakkaColIndex]).Value2?.ToString() ?? "";
                        if (!val.Contains("経済学科"))
                        {
                            wrongDataFound = true; // 経済学科以外が見えている＝不正解
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
                // - (オプション) 全行が表示されているわけではないこと（フィルタが効いている証拠）
                int totalRows = dataBody.Rows.Count;
                bool isFiltered = (visibleRows < totalRows); 

                // データセットによっては全件経済学科の可能性もあるため isFiltered は必須にしないが、通常は必須。
                // ここでは「間違ったデータが見えていない」ことを最優先。
                if (!wrongDataFound && correctDataFound)
                {
                    Console.WriteLine("[DEBUG] Task 4 Passed: Only '経済学科' is visible.");
                    return true;
                }

                Console.WriteLine("[DEBUG] Task 4 Failed.");
                return false;
            }, sheetNameOrNull: "担当者リスト"); // FindTableロジックのためにシート名を渡す設計に変更が必要だが、ここでは簡易化
        }

        // ==========================================
        // タスク5: 合計列追加・範囲変更 (修正版・残留チェック強化)
        // ==========================================
        private bool CheckTask_1_2_05_Impl(string filePath)
        {
            return ProcessTableTask(filePath, "イベント売上", (table) =>
            {
                // A. 範囲チェック (A4:G16)
                string rangeAddr = table.Range.Address.Replace("$", "").Replace(" ", "").ToUpper();
                if (!rangeAddr.EndsWith("G16") || table.Range.Columns.Count != 7)
                {
                    Console.WriteLine($"[DEBUG] Task 5 Failed: Range is {rangeAddr} (Expected ...G16).");
                    return false;
                }

                // B. 残留ゴミチェック (H列: H4:H16)
                // 範囲を一度広げてから戻すと、H列に「塗りつぶし」や「罫線」が残るためこれを検知する
                try
                {
                    Worksheet ws = table.Parent as Worksheet;
                    Range ghostRange = ws.Range["H4:H16"];

                    foreach (Range cell in ghostRange)
                    {
                        // チェック1: 背景色の残留 (Interior.ColorIndex)
                        // xlNone (-4142) でなければ、色が残っているとみなす
                        if ((int)cell.Interior.ColorIndex != -4142)
                        {
                            Console.WriteLine($"[DEBUG] Task 5 Failed: Residual fill color found at {cell.Address}.");
                            return false;
                        }

                        // チェック2: 罫線の残留
                        // 左辺(xlEdgeLeft)はテーブルと接しているため無視し、右・上・下のみチェックする
                        // xlEdgeRight=10, xlEdgeTop=8, xlEdgeBottom=9
                        int[] bordersToCheck = { 10, 8, 9 }; 
                        
                        foreach (int borderIndex in bordersToCheck)
                        {
                            // LineStyle が xlNone (-4142) でなければ線が残っている
                            // ※ cell.Borders[...] を使うとCOMエラーが出にくいため個別に取得
                            Border border = cell.Borders[(XlBordersIndex)borderIndex];
                            if ((int)border.LineStyle != -4142)
                            {
                                 Console.WriteLine($"[DEBUG] Task 5 Failed: Residual border found at {cell.Address} (Side: {borderIndex}).");
                                 return false;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[DEBUG] Error checking ghost range: {ex.Message}");
                    // チェック中にエラーが出た場合は安全策として不合格にする（厳密性優先）
                    return false;
                }

                Console.WriteLine("[DEBUG] Task 5 Passed: Range correct and clean.");
                return true;
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
                if (workbook == null) return false;

                // ワークシート特定 (Task4の場合、部分一致などで探すロジックが必要ならここに実装)
                foreach(Worksheet ws in workbook.Worksheets) {
                    if(ws.Name == targetSheet) { worksheet = ws; break; }
                }
                if (worksheet == null) return false;

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

                if (table == null) return false;

                // 実際の判定ロジックを実行
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
            try
            {
                Range data = table.DataBodyRange;
                if (data == null || data.Columns.Count < 2) return false;
                
                int lastCol = data.Columns.Count;
                int checkRow = 1;
                
                // 普通の列（中間の列、例：2列目や3列目）と最後の列を比較
                // 最初の列は特殊な場合があるので、中間の列を使う
                int normalCol = Math.Min(2, lastCol - 1); // 2列目、または最後から2列目
                
                Range normalCell = (Range)data.Cells[checkRow, normalCol];
                Range lastCell = (Range)data.Cells[checkRow, lastCol];
                
                // DisplayFormat.Interior.Color (見た目通りの色を取得する)
                double normalColColor = (double)normalCell.DisplayFormat.Interior.Color;
                double lastColColor = (double)lastCell.DisplayFormat.Interior.Color;
                
                // 太字チェックも追加
                dynamic lastFont = lastCell.Font;
                bool isBold = lastFont.Bold;

                // 色が違う OR 太字になっている
                return (normalColColor != lastColColor) || isBold;
            }
            catch { return false; }
        }

    }
}