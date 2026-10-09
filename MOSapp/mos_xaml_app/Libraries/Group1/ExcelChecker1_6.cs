using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Libraries;

namespace Libraries.Group1
{
    public class ExcelChecker1_6
    {
        private static readonly string TARGET_FILE_PATH =
            MOSExcelMogiApp.Infrastructure.DataPathHelper.GetWorkingFilePath(1, 6);

        private const string ExpectedHyperlinkAddress = "https://rabbitway.jp/service_mos";
        private const string ExpectedHyperlinkText = "パソコン資格講座のご相談";
        private const string ExpectedDocumentTitle = "売上一覧";

        // Public wrappers
        public bool CheckTask_1_6_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_6_01_Impl(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        public bool CheckTask_1_6_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_6_02_Impl(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        public bool CheckTask_1_6_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_6_03_Impl(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        public bool CheckTask_1_6_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(ExcelScoreExplanation.UnavailableText);
                return CheckTask_1_6_04_Impl(filePath);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        public string ValidateProject_6()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: project6.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: project6.xlsxが見つかりません。";
                }

                var results = new List<string>();
                bool task1 = CheckTask_1_6_01_Impl(TARGET_FILE_PATH);
                results.Add($"Task 6-1 (ウィンドウ枠の固定): {(task1 ? "OK" : "NG")}");

                bool task2 = CheckTask_1_6_02_Impl(TARGET_FILE_PATH);
                results.Add($"Task 6-2 (ハイパーリンク): {(task2 ? "OK" : "NG")}");

                bool task3 = CheckTask_1_6_03_Impl(TARGET_FILE_PATH);
                results.Add($"Task 6-3 (通貨書式): {(task3 ? "OK" : "NG")}");

                bool task4 = CheckTask_1_6_04_Impl(TARGET_FILE_PATH);
                results.Add($"Task 6-4 (プロパティタイトル): {(task4 ? "OK" : "NG")}");

                return string.Join("\n", results);
            }
            catch (Exception ex)
            {
                return $"エラー: {ex.Message}";
            }
        }

        // Helpers
        private bool IsExcelFileOpen(string filePath)
        {
            Application excelApp = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                foreach (Workbook workbook in excelApp.Workbooks)
                {
                    if (string.Equals(workbook.FullName, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);
            }
        }

        private string GetCurrentExcelFilePath()
        {
            Application excelApp = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
                return null;
            }
            catch (COMException)
            {
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
        // タスク6-1: ウィンドウ枠の固定（1～4行目）
        // ==========================================
        private bool CheckTask_1_6_01_Impl(string filePath)
        {
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
                    excelApp = new Application { Visible = false };
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                worksheet.Activate();
                Window window = excelApp.ActiveWindow;
                if (!window.FreezePanes)
                    return Miss("ウィンドウ枠が固定されていません。");

                int splitRow = window.SplitRow;
                if (splitRow >= 4)
                    return true;

                if (splitRow <= 0)
                    return Miss("1～4行目が常に表示されるようになっていません。");
                return Miss($"ウィンドウ枠の固定が{splitRow}行目までになっています。");
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        // ==========================================
        // タスク6-2: ハイパーリンク
        // ==========================================
        private bool CheckTask_1_6_02_Impl(string filePath)
        {
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
                    excelApp = new Application { Visible = false };
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Hyperlink best = null;
                int bestScore = -1;
                var candidates = new List<Hyperlink>();

                foreach (Hyperlink hyperlink in worksheet.Hyperlinks)
                {
                    candidates.Add(hyperlink);
                    string address = hyperlink.Address ?? "";
                    string text = hyperlink.TextToDisplay ?? "";
                    int score = 0;
                    if (address.IndexOf(ExpectedHyperlinkAddress, StringComparison.OrdinalIgnoreCase) >= 0)
                        score += 2;
                    if (text.Contains(ExpectedHyperlinkText))
                        score += 2;
                    if (score > bestScore)
                    {
                        bestScore = score;
                        best = hyperlink;
                    }
                }

                if (candidates.Count == 0)
                    return Miss("ハイパーリンクがありません。");

                if (best != null && bestScore >= 4)
                    return true;

                string bestAddress = best?.Address ?? "";
                string bestText = best?.TextToDisplay ?? "";
                bool addressOk = bestAddress.IndexOf(ExpectedHyperlinkAddress, StringComparison.OrdinalIgnoreCase) >= 0;
                bool textOk = bestText.Contains(ExpectedHyperlinkText);

                if (!addressOk)
                {
                    if (string.IsNullOrWhiteSpace(bestAddress))
                        ExcelScoreExplanation.Note("ハイパーリンクのリンク先が設定されていません。");
                    else
                        ExcelScoreExplanation.Note($"ハイパーリンクのリンク先が「{Quote(bestAddress)}」になっています。");
                }
                if (!textOk)
                {
                    if (string.IsNullOrWhiteSpace(bestText))
                        ExcelScoreExplanation.Note("ハイパーリンクの表示文字列が違います。");
                    else
                        ExcelScoreExplanation.Note($"ハイパーリンクの表示文字列が「{Quote(bestText)}」になっています。");
                }
                return false;
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        // ==========================================
        // タスク6-3: 通貨書式（小数点なし）
        // ==========================================
        private bool CheckTask_1_6_03_Impl(string filePath)
        {
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
                    excelApp = new Application { Visible = false };
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                worksheet = FindWorksheet(workbook, "販売実績");
                if (worksheet == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                Range targetRange = worksheet.Range["B5:G11"];
                return EvaluateCurrencyFormats(targetRange);
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private static bool EvaluateCurrencyFormats(Range targetRange)
        {
            var notCurrency = new List<string>();
            var hasDecimals = new List<string>();

            foreach (Range cell in targetRange.Cells)
            {
                string numberFormat = cell.NumberFormat as string ?? "";
                string address = CellAddress(cell);
                bool isCurrency = IsCurrencyFormat(numberFormat);
                bool noDecimals = !HasDecimalPlaces(numberFormat);

                if (!isCurrency)
                    notCurrency.Add(address);
                else if (!noDecimals)
                    hasDecimals.Add(address);
            }

            if (notCurrency.Count == 0 && hasDecimals.Count == 0)
                return true;

            if (notCurrency.Count > 0)
                ExcelScoreExplanation.Note($"{JoinNames(notCurrency)}が通貨の表示になっていません。");
            if (hasDecimals.Count > 0)
                ExcelScoreExplanation.Note($"{JoinNames(hasDecimals)}に小数点が表示されています。");
            return false;
        }

        private static bool IsCurrencyFormat(string numberFormat)
        {
            if (string.IsNullOrEmpty(numberFormat))
                return false;
            // 「通貨」書式は ¥ / $ / [$…] を含むことが多い
            return numberFormat.IndexOf('¥') >= 0
                || numberFormat.IndexOf('$') >= 0
                || numberFormat.IndexOf("[$", StringComparison.Ordinal) >= 0;
        }

        private static bool HasDecimalPlaces(string numberFormat)
        {
            if (string.IsNullOrEmpty(numberFormat))
                return false;
            // 小数点以下の桁を表す一般的パターン
            return numberFormat.Contains(".0") || numberFormat.Contains(".#") || numberFormat.Contains(".00");
        }

        // ==========================================
        // タスク6-4: プロパティのタイトル
        // 誤設定の検出は情報パネル初期表示に近い「タグ」「分類」に限定する。
        // ==========================================
        private bool CheckTask_1_6_04_Impl(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application { Visible = false };
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                    return Miss(ExcelScoreExplanation.UnavailableText);

                try
                {
                    dynamic properties = workbook.BuiltinDocumentProperties;
                    string title = TryGetBuiltinPropertyText(properties, "Title");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Task 6-4 Title: '{title}'");

                    if (!string.IsNullOrEmpty(title) && title.Contains(ExpectedDocumentTitle))
                        return true;

                    string keywords = TryGetBuiltinPropertyText(properties, "Keywords");
                    string category = TryGetBuiltinPropertyText(properties, "Category");

                    bool titleEmpty = string.IsNullOrWhiteSpace(title);
                    NoteWrongPropertyPlacement("タグ", keywords);
                    NoteWrongPropertyPlacement("分類", category);

                    bool keywordsNoted = PropertyLooksLikeExpectedAnswer(keywords);
                    bool categoryNoted = PropertyLooksLikeExpectedAnswer(category);

                    if (!titleEmpty)
                        ExcelScoreExplanation.Note($"プロパティのタイトルが「{Quote(title)}」になっています。");
                    else if (!keywordsNoted && !categoryNoted)
                        ExcelScoreExplanation.Note("プロパティのタイトルが設定されていません。");

                    return false;
                }
                catch
                {
                    return Miss(ExcelScoreExplanation.UnavailableText);
                }
            }
            catch (Exception)
            {
                return Miss(ExcelScoreExplanation.UnavailableText);
            }
        }

        /// <summary>
        /// タグ／分類へ誤配置したときの理由。正解文言なら場所のみ、それ以外は文言も出す。
        /// </summary>
        private static void NoteWrongPropertyPlacement(string propertyLabelJa, string value)
        {
            if (!PropertyLooksLikeExpectedAnswer(value))
                return;

            if (value.Contains(ExpectedDocumentTitle))
                ExcelScoreExplanation.Note($"プロパティの{propertyLabelJa}に設定されています。");
            else
                ExcelScoreExplanation.Note($"プロパティの{propertyLabelJa}に設定されていて、「{Quote(value)}」となっています。");
        }

        private static string TryGetBuiltinPropertyText(dynamic properties, string propertyName)
        {
            try
            {
                dynamic prop = properties[propertyName];
                object value = prop?.Value;
                return value?.ToString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>タイトル以外へ誤入力した可能性があるか（「売上一覧」「売上」など）。</summary>
        private static bool PropertyLooksLikeExpectedAnswer(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;
            return value.IndexOf("売上", StringComparison.Ordinal) >= 0;
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
    }
}
