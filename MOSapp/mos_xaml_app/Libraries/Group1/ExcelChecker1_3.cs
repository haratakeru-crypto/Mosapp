using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Libraries;

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

        // config tabs["1"] project 2 用エイリアス（採点は CheckTask_1_{projectId}_{task} を参照）
        public bool CheckTask_1_2_01() => CheckTask_1_3_01();
        public bool CheckTask_1_2_02() => CheckTask_1_3_02();
        public bool CheckTask_1_2_03() => CheckTask_1_3_03();
        public bool CheckTask_1_2_04() => CheckTask_1_3_04();
        public bool CheckTask_1_2_05() => CheckTask_1_3_05();
        public bool CheckTask_1_2_06() => CheckTask_1_3_06();
        public bool CheckTask_1_2_07() => CheckTask_1_3_07();

        // 共通エラーハンドリング
        private bool RunCheck(Func<string, bool> checkImpl, string taskName)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] {taskName} called");
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return checkImpl(filePath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in {taskName}: {ex.Message}");
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        // ==========================================
        // 実装メソッド
        // ==========================================

        // タスク3-1: セルスタイルの適用 (A11:G11 に「集計」スタイル)
        private bool CheckTask_1_3_01_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "下半期売上", (worksheet) =>
            {
                Range targetRange = worksheet.Range["A11:G11"];
                var missing = new List<string>();
                foreach (Range cell in targetRange.Cells)
                {
                    if (!IsTotalStyle(cell))
                        missing.Add(CellAddress(cell));
                }

                var extras = new List<string>();
                foreach (Range cell in worksheet.Range["A10:G10"].Cells)
                {
                    if (IsTotalStyle(cell))
                        extras.Add(CellAddress(cell));
                }
                foreach (Range cell in worksheet.Range["A12:G12"].Cells)
                {
                    if (IsTotalStyle(cell))
                        extras.Add(CellAddress(cell));
                }

                if (missing.Count == 0 && extras.Count == 0)
                    return true;
                if (missing.Count > 0)
                    ExcelScoreExplanation.Note($"{JoinNames(missing)}に「集計」のセルのスタイルが設定されていません。");
                if (extras.Count > 0)
                    ExcelScoreExplanation.Note($"{JoinNames(extras)}にも「集計」のセルのスタイルが設定されています。");
                return false;
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

                int topIndent = Convert.ToInt32(topNeighbor.IndentLevel);
                int bottomIndent = Convert.ToInt32(bottomNeighbor.IndentLevel);

                if (topIndent != 0)
                    ExcelScoreExplanation.Note($"B4のインデントが{topIndent}になっています。");
                if (bottomIndent != 0)
                    ExcelScoreExplanation.Note($"B45のインデントが{bottomIndent}になっています。");

                var wrong = new List<string>();
                int sampleIndent = -1;
                foreach (Range cell in targetRange)
                {
                    int indent = Convert.ToInt32(cell.IndentLevel);
                    if (indent != 2)
                    {
                        wrong.Add(CellAddress(cell));
                        if (sampleIndent < 0)
                            sampleIndent = indent;
                    }
                }
                if (wrong.Count > 0)
                    ExcelScoreExplanation.Note($"{JoinNames(wrong)}のインデントが{sampleIndent}になっています。");

                if (topIndent == 0 && bottomIndent == 0 && wrong.Count == 0)
                    return true;
                return false;
            });
        }

        // ==========================================
        // タスク3-3: 選択範囲内で中央 (修正版・定数訂正)
        // ==========================================
        private bool CheckTask_1_3_03_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "社員リスト", (ws) =>
            {
                Range targetCell = ws.Range["A2"];
                const int XlHAlignCenterAcrossSelection = 7;

                int align;
                try
                {
                    align = Convert.ToInt32(targetCell.HorizontalAlignment);
                }
                catch
                {
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }

                bool alignOk = align == XlHAlignCenterAcrossSelection;
                if (!alignOk)
                    ExcelScoreExplanation.Note($"A2の文字の配置が「{DescribeAlignment(align)}」になっています。");

                Range rightNeighbor = ws.Range["G2"];
                int rightAlign = 0;
                try
                {
                    rightAlign = Convert.ToInt32(rightNeighbor.HorizontalAlignment);
                }
                catch { }

                bool rightOk = rightAlign != XlHAlignCenterAcrossSelection;
                if (!rightOk)
                    ExcelScoreExplanation.Note("G2まで選択範囲内で中央になっています。");

                return alignOk && rightOk;
            });
        }

        // タスク3-4: 列幅保持コピー
        private bool CheckTask_1_3_04_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "担当者別売上", (worksheet) =>
            {
                Range sourceRange = worksheet.Range["H5:K19"];
                Range targetRange = worksheet.Range["A5:D19"];
                var mismatch = new List<string>();

                for (int j = 1; j <= 4; j++)
                {
                    Range sourceCol = (Range)sourceRange.Cells[1, j];
                    Range targetCol = (Range)targetRange.Cells[1, j];

                    double w1 = Convert.ToDouble(sourceCol.ColumnWidth);
                    double w2 = Convert.ToDouble(targetCol.ColumnWidth);

                    if (Math.Abs(w1 - w2) > 0.1)
                    {
                        string srcLetter = ColumnLetter(sourceCol.Column);
                        string dstLetter = ColumnLetter(targetCol.Column);
                        mismatch.Add($"{dstLetter}列（コピー元{srcLetter}列と幅が違う）");
                    }
                }

                if (mismatch.Count == 0)
                    return true;
                return Miss($"列幅を保持した貼り付けになっていません。{JoinNames(mismatch)}。");
            });
        }

        // タスク3-5: 取り消し線 (範囲厳格化)
        private bool CheckTask_1_3_05_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "業務予定", (ws) =>
            {
                Range targetRange = ws.Range["C5:C11"];

                Range topNeighbor = ws.Range["C4"];
                Range bottomNeighbor = ws.Range["C12"];

                bool topStrike = topNeighbor.Font.Strikethrough is bool bTop && bTop;
                bool bottomStrike = bottomNeighbor.Font.Strikethrough is bool bBot && bBot;

                var outside = new List<string>();
                if (topStrike) outside.Add("C4");
                if (bottomStrike) outside.Add("C12");
                if (outside.Count > 0)
                    ExcelScoreExplanation.Note($"{JoinNames(outside)}にも取り消し線が設定されています。");

                var missing = new List<string>();
                foreach (Range cell in targetRange)
                {
                    if (!(cell.Font.Strikethrough is bool b && b))
                        missing.Add(CellAddress(cell));
                }
                if (missing.Count > 0)
                    ExcelScoreExplanation.Note($"{JoinNames(missing)}に取り消し線が設定されていません。");

                return outside.Count == 0 && missing.Count == 0;
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
                bool isCenter = (align == -4108 || align == 7);

                if (!isMerged && !isCenter)
                    return true;
                if (isMerged)
                    ExcelScoreExplanation.Note("B2のセルの結合が解除されていません。");
                if (isCenter)
                    ExcelScoreExplanation.Note($"B2の文字の配置が「{DescribeAlignment(align)}」のままです。");
                return false;
            });
        }

        // ==========================================
        // タスク3-7: 行の削除 (修正版・正解件数20件)
        // ==========================================
        private bool CheckTask_1_3_07_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "参加者一覧", (ws) =>
            {
                int nameCol = -1;
                int headerRow = -1;

                for (int r = 1; r <= 10; r++)
                {
                    for (int c = 1; c <= 20; c++)
                    {
                        Range cell = (Range)ws.Cells[r, c];
                        string val = Convert.ToString(cell.Value2);

                        if (!string.IsNullOrEmpty(val) && val.Contains("氏名"))
                        {
                            nameCol = c;
                            headerRow = r;
                            break;
                        }
                    }
                    if (headerRow != -1) break;
                }

                if (nameCol == -1)
                {
                    nameCol = 2;
                    headerRow = 0;
                }

                var targets = new HashSet<string> { "風間健太郎", "平井元" };
                var remainingTargets = new List<string>();
                const int ExpectedCount = 20;

                int currentDataCount = 0;
                int scanLimit = 60;

                for (int r = headerRow + 1; r <= scanLimit; r++)
                {
                    Range cell = (Range)ws.Cells[r, nameCol];
                    string rawVal = Convert.ToString(cell.Value2);

                    if (string.IsNullOrWhiteSpace(rawVal)) continue;
                    if (rawVal.Contains("氏名")) continue;

                    string normalized = rawVal.Replace(" ", "")
                                              .Replace("　", "")
                                              .Replace("\u00A0", "")
                                              .Replace("\t", "");

                    if (targets.Contains(normalized))
                        remainingTargets.Add(Quote(rawVal.Trim()));

                    currentDataCount++;
                }

                if (remainingTargets.Count > 0)
                    ExcelScoreExplanation.Note($"削除対象の行（{JoinNames(remainingTargets)}）が残っています。");
                if (currentDataCount != ExpectedCount)
                    ExcelScoreExplanation.Note($"データの行数が{currentDataCount}行になっています。");
                if (remainingTargets.Count == 0 && currentDataCount == ExpectedCount)
                    return true;
                return false;
            });
        }

        // ==========================================
        // ヘルパーメソッド
        // ==========================================

        private static bool IsTotalStyle(Range cell)
        {
            try
            {
                dynamic style = cell.Style;
                string styleName = (string)(style.NameLocal ?? string.Empty);
                return styleName == "集計" || styleName == "Total";
            }
            catch
            {
                return false;
            }
        }

        private bool CheckTaskBasic(string filePath, string sheetName, Func<Worksheet, bool> checkLogic)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try { excelApp = (Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { return Miss(ExcelScoreExplanation.UnavailableText); }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                worksheet = FindWorksheet(workbook, sheetName);
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                return checkLogic(worksheet);
            }
            catch { return Miss(ExcelScoreExplanation.UnavailableText); }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
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

        private static string CellAddress(Range cell)
        {
            try { return cell.Address[false, false]; }
            catch { return "?"; }
        }

        private static string ColumnLetter(int column)
        {
            if (column <= 0)
                return "?";
            string result = "";
            int n = column;
            while (n > 0)
            {
                n--;
                result = (char)('A' + (n % 26)) + result;
                n /= 26;
            }
            return result;
        }

        private static string DescribeAlignment(int align)
        {
            switch (align)
            {
                case 7: return "選択範囲内で中央";
                case -4108: return "中央揃え";
                case -4131: return "左揃え";
                case -4152: return "右揃え";
                case -4130: return "均等割り付け";
                case 1: return "標準";
                default: return $"その他（{align}）";
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
