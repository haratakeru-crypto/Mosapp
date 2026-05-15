using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group1
{
    public class ExcelChecker1_3
    {
        // ==========================================
        // 公開メソッド（呼び出し元）
        // ==========================================

        public bool CheckTask_1_3_01() => RunCheck(CheckTask_1_3_01_Impl, "Task 3-1 (Style)");
        public bool CheckTask_1_3_02() => RunCheck(CheckTask_1_3_02_Impl, "Task 3-2 (Indent)");
        public bool CheckTask_1_3_03() => RunCheck(CheckTask_1_3_03_Impl, "Task 3-3 (Center)");
        public bool CheckTask_1_3_04() => RunCheck(CheckTask_1_3_04_Impl, "Task 3-4 (Copy Width)");
        public bool CheckTask_1_3_05() => RunCheck(CheckTask_1_3_05_Impl, "Task 3-5 (Strike)");
        public bool CheckTask_1_3_06() => RunCheck(CheckTask_1_3_06_Impl, "Task 3-6 (Unmerge)");
        public bool CheckTask_1_3_07() => RunCheck(CheckTask_1_3_07_Impl, "Task 3-7 (Delete Row)");

        // 共通エラーハンドリング
        private bool RunCheck(Func<string, bool> checkImpl, string taskName)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] {taskName} called");
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath)) return false;
                return checkImpl(filePath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in {taskName}: {ex.Message}");
                return false;
            }
        }

        // ==========================================
        // 実装メソッド
        // ==========================================

        // タスク3-1: セルスタイルの適用
        private bool CheckTask_1_3_01_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "下半期売上", (worksheet) =>
            {
                Range targetCell = worksheet.Range["A2"];
                dynamic style = targetCell.Style;
                string styleName = style.NameLocal; // "タイトル"
                string styleNameEng = style.Name;   // "Title"
                return styleName == "タイトル" || styleName == "Title";
            });
        }

        // タスク3-2: 左インデント2文字 (範囲厳格化)
        private bool CheckTask_1_3_02_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "社員リスト", (ws) =>
            {
                Range targetRange = ws.Range["B5:B44"];

                // 範囲外チェック (B4:見出し, B45:下)
                Range topNeighbor = ws.Range["B4"];
                Range bottomNeighbor = ws.Range["B45"];

                // 安全な型変換のために Convert.ToInt32 を使用
                int topIndent = Convert.ToInt32(topNeighbor.IndentLevel);
                int bottomIndent = Convert.ToInt32(bottomNeighbor.IndentLevel);

                if (topIndent != 0)
                {
                    Console.WriteLine("[DEBUG] Task 2 Failed: Header B4 is indented.");
                    return false;
                }
                if (bottomIndent != 0)
                {
                    Console.WriteLine("[DEBUG] Task 2 Failed: Bottom cell B45 is indented.");
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

        // ==========================================
        // タスク3-3: 選択範囲内で中央 (修正版・定数訂正)
        // ==========================================
        private bool CheckTask_1_3_03_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "社員リスト", (ws) =>
            {
                // 正解の範囲: A2:F2
                Range targetCell = ws.Range["A2"];
                
                // 【修正ポイント】
                // 選択範囲内で中央 (xlHAlignCenterAcrossSelection) の値は「7」です。
                // 以前のコードでは中央揃え(-4108)で判定していたため修正しました。
                const int XlHAlignCenterAcrossSelection = 7;

                // 型変換を安全に行う
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
                    Console.WriteLine($"[DEBUG] Task 3 Failed: Alignment is {align} (Expected 7).");
                    return false;
                }

                // 範囲外チェック (G2)
                // G2まで「選択範囲内で中央」になっていたら範囲選択ミス
                Range rightNeighbor = ws.Range["G2"];
                int rightAlign = 0;
                try
                {
                    rightAlign = Convert.ToInt32(rightNeighbor.HorizontalAlignment);
                }
                catch { } // 無視

                if (rightAlign == XlHAlignCenterAcrossSelection)
                {
                    Console.WriteLine("[DEBUG] Task 3 Failed: Range extends too far (G2 is included).");
                    return false;
                }

                Console.WriteLine("[DEBUG] Task 3 Passed.");
                return true;
            });
        }

        // タスク3-4: 列幅保持コピー
        private bool CheckTask_1_3_04_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "担当者別売上", (worksheet) =>
            {
                Range sourceRange = worksheet.Range["H5:K19"];
                Range targetRange = worksheet.Range["A5:D19"];

                // 列幅チェック
                for (int j = 1; j <= 4; j++)
                {
                    Range sourceCol = (Range)sourceRange.Cells[1, j];
                    Range targetCol = (Range)targetRange.Cells[1, j];
                    
                    double w1 = Convert.ToDouble(sourceCol.ColumnWidth);
                    double w2 = Convert.ToDouble(targetCol.ColumnWidth);

                    if (Math.Abs(w1 - w2) > 0.1) return false;
                }
                return true;
            });
        }

        // タスク3-5: 取り消し線 (範囲厳格化)
        private bool CheckTask_1_3_05_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "業務予定", (ws) =>
            {
                Range targetRange = ws.Range["C5:C11"];
                
                // 範囲外チェック (C4, C12)
                Range topNeighbor = ws.Range["C4"];
                Range bottomNeighbor = ws.Range["C12"];
                
                // Null安全な変換
                bool topStrike = topNeighbor.Font.Strikethrough is bool bTop && bTop;
                bool bottomStrike = bottomNeighbor.Font.Strikethrough is bool bBot && bBot;

                if (topStrike || bottomStrike)
                {
                    Console.WriteLine("[DEBUG] Task 5 Failed: Range incorrect.");
                    return false;
                }

                foreach (Range cell in targetRange)
                {
                    if (!(cell.Font.Strikethrough is bool b && b)) return false;
                }
                return true;
            });
        }

        // タスク3-6: 結合解除
        private bool CheckTask_1_3_06_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "参加者一覧", (worksheet) =>
            {
                Range targetCell = worksheet.Range["B2"];
                bool isMerged = (bool)targetCell.MergeCells;
                int align = Convert.ToInt32(targetCell.HorizontalAlignment);
                // 中央揃え(-4108)や選択範囲内中央(7)が解除されているか
                bool isCenter = (align == -4108 || align == 7);
                return !isMerged && !isCenter;
            });
        }

        // ==========================================
        // タスク3-7: 行の削除 (修正版・正解件数20件)
        // ==========================================
        private bool CheckTask_1_3_07_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "参加者一覧", (ws) =>
            {
                // 1. 「氏名」列とヘッダー行を特定
                int nameCol = -1;
                int headerRow = -1;

                // ヘッダー検索: 1～10行目を確認
                for (int r = 1; r <= 10; r++) 
                {
                    for (int c = 1; c <= 20; c++) 
                    {
                        Range cell = (Range)ws.Cells[r, c];
                        string val = Convert.ToString(cell.Value2);
                        
                        // "氏名" を含むセルをヘッダーとみなす
                        if (!string.IsNullOrEmpty(val) && val.Contains("氏名"))
                        {
                            nameCol = c;
                            headerRow = r;
                            break;
                        }
                    }
                    if (headerRow != -1) break;
                }

                // 見つからない場合のデフォルト (B列, 行なし)
                if (nameCol == -1) 
                {
                    nameCol = 2;
                    headerRow = 0;
                }

                // 2. 削除対象 (この名前が残っていたら即NG)
                var targets = new HashSet<string> { "風間健太郎", "平井元" };

                // 3. データ件数チェック
                // 純粋なデータの数が 20行 であれば正解
                const int ExpectedCount = 20;
                
                int currentDataCount = 0;
                int scanLimit = 60; // データ範囲の余裕を見て60行目までスキャン

                // ヘッダー行の次の行からカウント開始
                for (int r = headerRow + 1; r <= scanLimit; r++)
                {
                    Range cell = (Range)ws.Cells[r, nameCol];
                    string rawVal = Convert.ToString(cell.Value2);
                    
                    // 空白セルはデータとして数えない
                    if (string.IsNullOrWhiteSpace(rawVal)) continue;

                    // (安全策) ヘッダー文字列そのものが重複して登場した場合も数えない
                    if (rawVal.Contains("氏名")) continue;
                    // (安全策) "タイトル"のような行が混ざっている場合も除外検討が必要だが
                    // 基本は「氏名」より下の非空白行をデータとみなす

                    // データの正規化（スペース除去）
                    string normalized = rawVal.Replace(" ", "")
                                              .Replace("　", "")
                                              .Replace("\u00A0", "")
                                              .Replace("\t", "");

                    // チェックA: 削除対象が残っていないか
                    if (targets.Contains(normalized))
                    {
                        Console.WriteLine($"[DEBUG] Task 7 Failed: Target '{rawVal}' still exists.");
                        return false; 
                    }

                    // 有効なデータとしてカウント
                    currentDataCount++;
                }

                // チェックB: データ件数が 20件 か
                if (currentDataCount != ExpectedCount)
                {
                    Console.WriteLine($"[DEBUG] Task 7 Failed: Count is {currentDataCount} (Expected {ExpectedCount}).");
                    return false;
                }

                Console.WriteLine($"[DEBUG] Task 7 Passed: Count is {ExpectedCount}.");
                return true;
            });
        }

        // ==========================================
        // ヘルパーメソッド
        // ==========================================

        private bool CheckTaskBasic(string filePath, string sheetName, Func<Worksheet, bool> checkLogic)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { return false; } // Excelが開いていない

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

        private Worksheet FindWorksheet(Workbook workbook, string sheetName)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return sheet;
            }
            return null;
        }

        private string GetCurrentExcelFilePath()
        {
            try
            {
                Application excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.ActiveWorkbook != null) return excelApp.ActiveWorkbook.FullName;
                return null;
            }
            catch { return null; }
        }
    }
}
