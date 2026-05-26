using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group1
{
    public class ExcelChecker1_9
    {
        // ラッパーメソッド
        public bool CheckTask_1_9_01() => RunCheck(CheckTask_1_9_01_Impl, "Task 9-1 (Display Formulas)");
        public bool CheckTask_1_9_02() => RunCheck(CheckTask_1_9_02_Impl, "Task 9-2 (Sort Data)");
        public bool CheckTask_1_9_03() => RunCheck(CheckTask_1_9_03_Impl, "Task 9-3 (Icon Set)");
        public bool CheckTask_1_9_04() => RunCheck(CheckTask_1_9_04_Impl, "Task 9-4 (Cond Format > 3000)");
        public bool CheckTask_1_9_05() => RunCheck(CheckTask_1_9_05_Impl, "Task 9-5 (Header Date)");
        public bool CheckTask_1_9_06() => RunCheck(CheckTask_1_9_06_Impl, "Task 9-6 (Footer Page/Total)");
        public bool CheckTask_1_9_07() => RunCheck(CheckTask_1_9_07_Impl, "Task 9-7 (Accessibility H7)");

        // 共通エラーハンドリング
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
        // タスク9-1: 数式の表示
        // ==========================================
        private bool CheckTask_1_9_01_Impl(string filePath)
        {
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
                        
                        if (ws.Name == "売上報告")
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

        // ==========================================
        // タスク9-2: 並べ替え
        // シート：「受注明細」 商品ID(昇順) -> 金額(降順)
        // ==========================================
        private bool CheckTask_1_9_02_Impl(string filePath)
        {
            return ProcessSheet(filePath, "受注明細", (ws) =>
            {
                Range usedRange = ws.UsedRange;
                object[,] values = (object[,])usedRange.Value2;
                if (values == null) return false;

                int rowCount = values.GetLength(0);
                int colCount = values.GetLength(1);

                // ヘッダー行を探す
                int headerRow = -1;
                int colID = -1;
                int colAmount = -1;

                for (int r = 1; r <= Math.Min(10, rowCount); r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        string val = Convert.ToString(values[r, c]);
                        if (val == "商品ID") colID = c;
                        if (val == "金額") colAmount = c;
                    }
                    if (colID != -1 && colAmount != -1)
                    {
                        headerRow = r;
                        break;
                    }
                }

                if (headerRow == -1) return false;

                // データの並び順チェック
                for (int r = headerRow + 2; r <= rowCount; r++)
                {
                    // 数値として比較を試みる（文字列の場合もあるのでTryParse）
                    string sIdCur = Convert.ToString(values[r, colID]);
                    string sIdPrev = Convert.ToString(values[r - 1, colID]);
                    
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

                    // 商品ID (昇順): 前の行 <= 今の行
                    // compareID < 0 なら "今の行 < 前の行" なのでNG
                    if (compareID < 0) 
                    {
                        Console.WriteLine($"[DEBUG] Sort Error Row {r}: ID {sIdCur} < {sIdPrev}");
                        return false;
                    }

                    // IDが同じ場合、金額 (降順): 前の行 >= 今の行
                    if (compareID == 0)
                    {
                        double amCur = 0, amPrev = 0;
                        try { amCur = Convert.ToDouble(values[r, colAmount]); } catch { }
                        try { amPrev = Convert.ToDouble(values[r - 1, colAmount]); } catch { }

                        // 降順なので、今の行が前の行より大きかったらNG
                        if (amCur > amPrev)
                        {
                            Console.WriteLine($"[DEBUG] Sort Error Row {r}: Amount {amCur} > {amPrev}");
                            return false;
                        }
                    }
                }
                return true;
            });
        }

        // ==========================================
        // タスク9-3: アイコンセット
        // シート：「下半期売上」 7月-12月 (G列～L列と推測されますが、UsedRange全体から探します)
        // 条件：3つの矢印（色分け） -> ID = 1
        // ==========================================
        private bool CheckTask_1_9_03_Impl(string filePath)
        {
            return ProcessSheet(filePath, "下半期売上", (ws) =>
            {
                dynamic usedRange = ws.UsedRange;
                dynamic formatConditions = usedRange.FormatConditions;

                Console.WriteLine($"[DEBUG] Task 9-3: Checking {formatConditions.Count} format conditions...");

                foreach (dynamic fc in formatConditions)
                {
                    try
                    {
                        int fcType = (int)fc.Type;
                        Console.WriteLine($"[DEBUG] Task 9-3: FormatCondition Type = {fcType}");
                        
                        // Type 6 = xlIconSet (アイコンセット)
                        if (fcType == 6) 
                        {
                            dynamic iconSet = fc.IconSet;
                            int setId = (int)iconSet.ID;
                            
                            Console.WriteLine($"[DEBUG] Task 9-3: Found IconSet with ID = {setId}");

                            // ID 1 が「3つの矢印（色分け）」です
                            if (setId == 1)
                            {
                                Console.WriteLine("[DEBUG] Task 9-3: Passed - IconSet ID = 1 found");
                                return true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // COMオブジェクトの操作でエラーが出る場合があるため無視して次へ
                        Console.WriteLine($"[DEBUG] Task 9-3: Warning - {ex.Message}");
                    }
                }

                Console.WriteLine("[DEBUG] Task 9-3: Failed - No matching IconSet found");
                return false;
            });
        }

        // ==========================================
        // タスク9-4: 条件付き書式 (>3000)
        // ==========================================
        private bool CheckTask_1_9_04_Impl(string filePath)
        {
            return ProcessSheet(filePath, "下半期売上", (ws) =>
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
                            
                            if (f1 == "=3000" || f1 == "3000")
                            {
                                Console.WriteLine("[DEBUG] Task 9-4 Passed.");
                                return true;
                            }
                        }
                    }
                    catch { }
                }
                return false;
            });
        }

        // ==========================================
        // タスク9-5: ヘッダー
        // ==========================================
        private bool CheckTask_1_9_05_Impl(string filePath)
        {
            return ProcessSheet(filePath, "受注明細", (ws) =>
            {
                string rightHeader = ws.PageSetup.RightHeader;
                return !string.IsNullOrEmpty(rightHeader) && rightHeader.Contains("&D");
            });
        }

        // ==========================================
        // タスク9-6: フッター
        // ==========================================
        private bool CheckTask_1_9_06_Impl(string filePath)
        {
            return ProcessSheet(filePath, "受注明細", (ws) =>
            {
                string rightFooter = ws.PageSetup.RightFooter;
                return !string.IsNullOrEmpty(rightFooter) && rightFooter.Contains("&P") && rightFooter.Contains("&N");
            });
        }

        // ==========================================
        // タスク9-7: アクセシビリティ (H7)
        // 条件：負の数が「黒いマイナス記号」で表示されること
        // ==========================================
        private bool CheckTask_1_9_07_Impl(string filePath)
        {
            return ProcessSheet(filePath, "下半期売上", (ws) =>
            {
                Range cell = ws.Range["H7"];
                
                // 表示形式を取得
                string numFormat = (string)cell.NumberFormatLocal;
                
                Console.WriteLine($"[DEBUG] Task 9-7: NumberFormat = '{numFormat}'");
                
                // Excelの表示形式の仕組み:
                // - "0" → 負の数は黒いマイナスで表示される → OK
                // - "[Red]0" → 負の数は赤色で表示される → NG
                // - "0;-0" → 負の数は黒いマイナスで表示される → OK
                // - "0;[Red]-0" → 負の数は赤色で表示される → NG
                // 
                // アクセシビリティチェックの要件:
                // 赤色指定がなければ、負の数は黒いマイナス記号で表示される
                
                bool hasRedFormat = numFormat.Contains("[Red]") || numFormat.Contains("[赤]");
                
                if (hasRedFormat)
                {
                    Console.WriteLine("[DEBUG] Task 9-7: Failed - Red format found in NumberFormat");
                    return false;
                }
                
                // 赤色指定がなければOK
                Console.WriteLine("[DEBUG] Task 9-7: Passed - No red format, black minus will be displayed");
                return true;
            });
        }

        // ==========================================
        // ヘルパーメソッド
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

                workbook = GetWorkbook(excelApp, filePath);
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
                Console.WriteLine($"[DEBUG] Error: {ex.Message}");
                return false;
            }
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

        private string GetCurrentExcelFilePath()
        {
            try
            {
                Application excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                return excelApp.ActiveWorkbook?.FullName;
            }
            catch { return null; }
        }
    }
}
