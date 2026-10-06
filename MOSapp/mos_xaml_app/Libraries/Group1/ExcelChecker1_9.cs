using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Libraries;

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

        // config tabs["1"] project 7 用エイリアス
        public bool CheckTask_1_7_01() => CheckTask_1_9_01();
        public bool CheckTask_1_7_02() => CheckTask_1_9_02();
        public bool CheckTask_1_7_03() => CheckTask_1_9_03();
        public bool CheckTask_1_7_04() => CheckTask_1_9_04();
        public bool CheckTask_1_7_05() => CheckTask_1_9_05();
        public bool CheckTask_1_7_06() => CheckTask_1_9_06();
        public bool CheckTask_1_7_07() => CheckTask_1_9_07();

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
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                try { originalSheet = excelApp.ActiveSheet as Worksheet; }
                catch { }

                bool targetSheetCorrect = false;
                var formulaOthers = new List<string>();

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
                        else if (isDisplayingFormulas)
                        {
                            formulaOthers.Add(ws.Name);
                        }
                    }
                    catch { }
                }

                if (originalSheet != null)
                {
                    try { originalSheet.Activate(); }
                    catch { }
                }

                if (targetSheetCorrect && formulaOthers.Count == 0)
                    return true;

                if (!targetSheetCorrect)
                    ExcelScoreExplanation.Note("シート「売上報告」で数式が表示されていません。");
                if (formulaOthers.Count > 0)
                    ExcelScoreExplanation.Note($"シート「{JoinNames(formulaOthers)}」でも数式が表示されています。");
                return false;
            }
            catch
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        // ==========================================
        // タスク9-2: 並べ替え
        // ==========================================
        private bool CheckTask_1_9_02_Impl(string filePath)
        {
            return ProcessSheet(filePath, "受注明細", (ws) =>
            {
                Range usedRange = ws.UsedRange;
                object[,] values = (object[,])usedRange.Value2;
                if (values == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                int rowCount = values.GetLength(0);
                int colCount = values.GetLength(1);

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

                if (headerRow == -1)
                    return Miss("「商品ID」または「金額」の列が見つかりません。");

                for (int r = headerRow + 2; r <= rowCount; r++)
                {
                    string sIdCur = Convert.ToString(values[r, colID]);
                    string sIdPrev = Convert.ToString(values[r - 1, colID]);

                    double dIdCur = 0, dIdPrev = 0;
                    bool isNum = double.TryParse(sIdCur, out dIdCur) && double.TryParse(sIdPrev, out dIdPrev);

                    int compareID = isNum
                        ? dIdCur.CompareTo(dIdPrev)
                        : string.Compare(sIdCur, sIdPrev, StringComparison.Ordinal);

                    if (compareID < 0)
                    {
                        return Miss($"「商品ID」が昇順になっていません（「{Quote(sIdPrev)}」のあとに「{Quote(sIdCur)}」があります）。");
                    }

                    if (compareID == 0)
                    {
                        double amCur = 0, amPrev = 0;
                        try { amCur = Convert.ToDouble(values[r, colAmount]); } catch { }
                        try { amPrev = Convert.ToDouble(values[r - 1, colAmount]); } catch { }

                        if (amCur > amPrev)
                        {
                            return Miss($"同じ「商品ID」内で「金額」が降順になっていません（{amPrev} のあとに {amCur} があります）。");
                        }
                    }
                }
                return true;
            });
        }

        // ==========================================
        // タスク9-3: アイコンセット（種類 + 範囲 D5:I12）
        // ==========================================
        private bool CheckTask_1_9_03_Impl(string filePath)
        {
            return ProcessSheet(filePath, "下半期売上", (ws) =>
            {
                const string expectedRange = "D5:I12";
                const int expectedIconId = 1; // 3つの矢印（色分け）
                dynamic usedRange = ws.UsedRange;
                dynamic formatConditions = usedRange.FormatConditions;

                Console.WriteLine($"[DEBUG] Task 9-3: Checking {formatConditions.Count} format conditions...");

                // 失敗理由用: 期待範囲のルールを最優先。無ければ他範囲の候補を1件。
                int? onExpectedIconId = null;
                string onExpectedApplies = null;
                int? otherIconId = null;
                string otherApplies = null;
                bool foundAnyIconSet = false;

                foreach (dynamic fc in formatConditions)
                {
                    try
                    {
                        int fcType = (int)fc.Type;
                        if (fcType != 6)
                            continue;

                        foundAnyIconSet = true;
                        dynamic iconSet = fc.IconSet;
                        int setId = (int)iconSet.ID;
                        string applies = TryGetAppliesToAddress(fc);
                        bool rangeOk = IsSameRangeAddress(applies, expectedRange);
                        Console.WriteLine($"[DEBUG] Task 9-3: IconSet ID={setId}, AppliesTo={applies}");

                        if (setId == expectedIconId && rangeOk)
                        {
                            Console.WriteLine("[DEBUG] Task 9-3: Passed - IconSet ID=1 on D5:I12");
                            return true;
                        }

                        if (rangeOk)
                        {
                            // 期待範囲に複数ある場合は後勝ち（より新しい設定を優先）
                            onExpectedIconId = setId;
                            onExpectedApplies = applies;
                        }
                        else if (otherIconId == null)
                        {
                            otherIconId = setId;
                            otherApplies = applies;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[DEBUG] Task 9-3: Warning - {ex.Message}");
                    }
                }

                if (!foundAnyIconSet)
                    return Miss("アイコンセットが設定されていません。");

                // 診断は期待範囲 D5:I12 を最優先
                int reportIconId;
                string reportApplies;
                if (onExpectedIconId.HasValue)
                {
                    reportIconId = onExpectedIconId.Value;
                    reportApplies = onExpectedApplies;
                }
                else if (otherIconId.HasValue)
                {
                    reportIconId = otherIconId.Value;
                    reportApplies = otherApplies;
                }
                else
                {
                    return Miss("アイコンセットが指定どおりではありません。");
                }

                bool iconOk = reportIconId == expectedIconId;
                bool rangeOkReport = IsSameRangeAddress(reportApplies, expectedRange);
                string iconName = DescribeIconSet(reportIconId);
                string rangeText = string.IsNullOrEmpty(reportApplies) ? "（不明）" : reportApplies;

                if (iconOk && !rangeOkReport)
                    return Miss($"アイコンセット「{iconName}」の範囲が「{Quote(rangeText)}」になっています。");

                if (!iconOk && rangeOkReport)
                    return Miss($"アイコンセットが「{iconName}」になっています。");

                if (!iconOk && !rangeOkReport)
                    return Miss($"アイコンセットが「{iconName}」、範囲が「{Quote(rangeText)}」になっています。");

                return Miss("アイコンセットが指定どおりではありません。");
            });
        }

        // ==========================================
        // タスク9-4: 条件付き書式 (>3000 + 濃い黄色の文字・黄色の背景 + D5:I12)
        // ==========================================
        private bool CheckTask_1_9_04_Impl(string filePath)
        {
            return ProcessSheet(filePath, "下半期売上", (ws) =>
            {
                const string expectedRange = "D5:I12";

                dynamic usedRange = ws.UsedRange;
                dynamic formatConditions = usedRange.FormatConditions;

                bool foundGreater = false;

                // 期待範囲のルールを優先して診断（後から見つかったものを採用）
                bool hasOnExpected = false;
                string onExpectedFormula = null;
                string onExpectedApplies = null;
                bool onExpectedValueOk = false;
                bool onExpectedFillOk = false;
                bool onExpectedFontOk = false;

                bool hasOther = false;
                string otherFormula = null;
                string otherApplies = null;
                bool otherValueOk = false;
                bool otherFillOk = false;
                bool otherFontOk = false;

                foreach (dynamic fc in formatConditions)
                {
                    try
                    {
                        if ((int)fc.Type != 1 || (int)fc.Operator != 5)
                            continue;

                        foundGreater = true;
                        string f1 = "";
                        try { f1 = Convert.ToString(fc.Formula1) ?? ""; } catch { }
                        f1 = (f1 ?? "").Trim();
                        string applies = TryGetAppliesToAddress(fc);
                        bool rangeOk = IsSameRangeAddress(applies, expectedRange);
                        bool valueOk = IsGreaterThan3000Formula(f1);

                        double fillColor = TryGetOleColor(() => fc.Interior.Color);
                        double fontColor = TryGetOleColor(() => fc.Font.Color);
                        bool fillOk = IsYellowFillPreset(fillColor);
                        bool fontOk = IsDarkYellowFontPreset(fontColor);

                        Console.WriteLine(
                            $"[DEBUG] Task 9-4: Formula1={f1}, AppliesTo={applies}, " +
                            $"Fill={DescribeOleColor(fillColor)}, Font={DescribeOleColor(fontColor)}, " +
                            $"valueOk={valueOk}, rangeOk={rangeOk}, fillOk={fillOk}, fontOk={fontOk}");

                        if (valueOk && rangeOk && fillOk && fontOk)
                        {
                            Console.WriteLine("[DEBUG] Task 9-4 Passed.");
                            return true;
                        }

                        if (rangeOk)
                        {
                            hasOnExpected = true;
                            onExpectedFormula = f1;
                            onExpectedApplies = applies;
                            onExpectedValueOk = valueOk;
                            onExpectedFillOk = fillOk;
                            onExpectedFontOk = fontOk;
                        }
                        else if (!hasOther)
                        {
                            hasOther = true;
                            otherFormula = f1;
                            otherApplies = applies;
                            otherValueOk = valueOk;
                            otherFillOk = fillOk;
                            otherFontOk = fontOk;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[DEBUG] Task 9-4: Warning - {ex.Message}");
                    }
                }

                if (!foundGreater)
                    return Miss("「指定の値より大きい」条件付き書式が設定されていません。");

                // 範囲・条件・書式を独立に理由化（ずれている項目だけ複数行）
                bool reportRangeOk = hasOnExpected;
                string reportFormula = hasOnExpected ? onExpectedFormula : otherFormula;
                string reportApplies = hasOnExpected ? onExpectedApplies : otherApplies;
                bool reportValueOk = hasOnExpected ? onExpectedValueOk : otherValueOk;
                bool reportFillOk = hasOnExpected ? onExpectedFillOk : otherFillOk;
                bool reportFontOk = hasOnExpected ? onExpectedFontOk : otherFontOk;
                bool reportFormatOk = reportFillOk && reportFontOk;

                string rangeText = string.IsNullOrEmpty(reportApplies) ? "（不明）" : reportApplies;
                string formulaText = string.IsNullOrEmpty(reportFormula) ? "（空）" : reportFormula;

                var reasons = new List<string>();
                if (!reportValueOk)
                    reasons.Add($"条件付き書式の値が「{Quote(formulaText)}」になっています。");
                if (!reportRangeOk)
                    reasons.Add($"条件付き書式の範囲が「{Quote(rangeText)}」になっています。");
                if (!reportFormatOk)
                    reasons.Add("書式が「濃い黄色の文字、黄色の背景」になっていません。");

                if (reasons.Count == 0)
                    return Miss("条件付き書式が指定どおりではありません。");

                foreach (string reason in reasons)
                    ExcelScoreExplanation.Note(reason);
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
                string rightHeader = ws.PageSetup.RightHeader ?? "";
                if (!string.IsNullOrEmpty(rightHeader) && rightHeader.Contains("&D"))
                    return true;

                if (string.IsNullOrWhiteSpace(rightHeader))
                    return Miss("ヘッダーの右に現在の日付がありません。");
                return Miss($"ヘッダーの右が「{Quote(DescribeHeaderFooter(rightHeader))}」になっています。");
            });
        }

        // ==========================================
        // タスク9-6: フッター（右に &P/&N）
        // ==========================================
        private bool CheckTask_1_9_06_Impl(string filePath)
        {
            return ProcessSheet(filePath, "受注明細", (ws) =>
            {
                string rightFooter = ws.PageSetup.RightFooter ?? "";
                bool hasP = rightFooter.Contains("&P");
                bool hasN = rightFooter.Contains("&N");
                bool hasSlashBetween = rightFooter.Contains("&P/&N");
                if (!string.IsNullOrEmpty(rightFooter) && hasSlashBetween)
                    return true;

                if (string.IsNullOrWhiteSpace(rightFooter))
                    return Miss("フッターの右に「P/N」がありません。");

                if (!hasP)
                    ExcelScoreExplanation.Note("フッターの右にページ番号がありません。");
                if (!hasN)
                    ExcelScoreExplanation.Note("フッターの右に総ページ数がありません。");
                if (hasP && hasN && !hasSlashBetween)
                    ExcelScoreExplanation.Note($"フッターの右が「{Quote(DescribeHeaderFooter(rightFooter))}」になっています。");
                return false;
            });
        }

        // ==========================================
        // タスク9-7: アクセシビリティ (H7)
        // ==========================================
        private bool CheckTask_1_9_07_Impl(string filePath)
        {
            return ProcessSheet(filePath, "下半期売上", (ws) =>
            {
                Range cell = ws.Range["H7"];
                string numFormat = (string)cell.NumberFormatLocal;
                Console.WriteLine($"[DEBUG] Task 9-7: NumberFormat = '{numFormat}'");

                bool hasRedFormat = !string.IsNullOrEmpty(numFormat)
                    && (numFormat.Contains("[Red]") || numFormat.Contains("[赤]"));

                if (hasRedFormat)
                    return Miss("H7の表示形式で負の数が赤色になっています。");

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
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = ws;
                        break;
                    }
                }
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                return checkLogic(worksheet);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error: {ex.Message}");
                return Miss(ExcelScoreExplanation.UnavailableText);
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

        private static string DescribeIconSet(int iconSetId)
        {
            switch (iconSetId)
            {
                case 1: return "3つの矢印（色分け）";
                case 2: return "3つの矢印（グレー）";
                case 3: return "3つの旗";
                case 4: return "3つの信号（枠なし）";
                case 5: return "3つの信号（枠あり）";
                case 6: return "3つの記号（丸）";
                case 7: return "3つの記号";
                case 8: return "4つの矢印（色分け）";
                case 9: return "4つの矢印（グレー）";
                case 10: return "赤から黒へ（4）";
                case 11: return "評価（4）";
                case 12: return "4つの信号";
                case 13: return "5つの矢印（色分け）";
                case 14: return "5つの矢印（グレー）";
                case 15: return "評価（5）";
                case 16: return "5つの四半分";
                case 17: return "3つの星";
                case 18: return "3つの三角";
                case 19: return "5つの枠";
                case 20: return "5つの評価";
                default: return $"別のアイコンセット（ID:{iconSetId}）";
            }
        }

        private static string TryGetAppliesToAddress(dynamic formatCondition)
        {
            try
            {
                Range applies = formatCondition.AppliesTo as Range;
                if (applies == null)
                    applies = formatCondition.AppliesTo;
                if (applies == null)
                    return null;
                return applies.Address[false, false];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] TryGetAppliesToAddress: {ex.Message}");
                return null;
            }
        }

        private static double TryGetOleColor(Func<object> getter)
        {
            try
            {
                object value = getter();
                if (value == null)
                    return double.NaN;
                return Convert.ToDouble(value);
            }
            catch
            {
                return double.NaN;
            }
        }

        private static bool TryUnpackRgb(double oleColor, out int r, out int g, out int b)
        {
            r = g = b = 0;
            if (double.IsNaN(oleColor))
                return false;
            int color = unchecked((int)oleColor);
            r = color & 0xFF;
            g = (color >> 8) & 0xFF;
            b = (color >> 16) & 0xFF;
            return true;
        }

        private static bool IsCloseRgb(double oleColor, int expectedR, int expectedG, int expectedB, int tolerance = 25)
        {
            if (!TryUnpackRgb(oleColor, out int r, out int g, out int b))
                return false;
            return Math.Abs(r - expectedR) <= tolerance
                && Math.Abs(g - expectedG) <= tolerance
                && Math.Abs(b - expectedB) <= tolerance;
        }

        /// <summary>Excelプリセット「黄色の背景」(RGB 255,235,156)。赤系塗りつぶしを除外。</summary>
        private static bool IsYellowFillPreset(double oleColor)
        {
            if (!IsCloseRgb(oleColor, 255, 235, 156, 25))
                return false;
            if (!TryUnpackRgb(oleColor, out int r, out int g, out int b))
                return false;
            // 赤背景(例: 255,199,206)は G≈B。黄は G が B より十分大きい。
            if (g - b < 50)
                return false;
            return r >= 230 && g >= 210;
        }

        /// <summary>Excelプリセット「濃い黄色の文字」(RGB 156,101,0)。濃い赤文字を除外。</summary>
        private static bool IsDarkYellowFontPreset(double oleColor)
        {
            if (!IsCloseRgb(oleColor, 156, 101, 0, 25))
                return false;
            if (!TryUnpackRgb(oleColor, out _, out int g, out int b))
                return false;
            // 濃い赤文字(例: 156,0,6)は G がほぼ0。
            if (g < 60)
                return false;
            if (b > 40)
                return false;
            return true;
        }

        private static bool IsGreaterThan3000Formula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return false;
            string text = formula.Trim().TrimStart('=');
            return text == "3000";
        }

        private static string DescribeOleColor(double oleColor)
        {
            if (!TryUnpackRgb(oleColor, out int r, out int g, out int b))
                return "（不明）";
            return $"RGB({r},{g},{b})";
        }

        private static bool IsSameRangeAddress(string actual, string expected)
        {
            string a = NormalizeRangeAddress(actual);
            string e = NormalizeRangeAddress(expected);
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(e))
                return false;
            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeRangeAddress(string address)
        {
            if (string.IsNullOrEmpty(address))
                return "";
            string text = address.Replace("$", "").Replace(" ", "").ToUpperInvariant();
            int bang = text.LastIndexOf('!');
            if (bang >= 0 && bang < text.Length - 1)
                text = text.Substring(bang + 1);
            return text;
        }

        private static string DescribeHeaderFooter(string code)
        {
            if (string.IsNullOrEmpty(code))
                return "（空）";
            return code
                .Replace("&D", "［日付］")
                .Replace("&P", "［ページ番号］")
                .Replace("&N", "［ページ数］")
                .Replace("&A", "［シート名］")
                .Replace("&F", "［ファイル名］")
                .Replace("&Z", "［パス］")
                .Replace("&T", "［時刻］");
        }
    }
}
