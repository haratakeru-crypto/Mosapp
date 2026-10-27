using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Libraries
{
    internal static class VariantCheckSupport
    {
        public static readonly string DiagLogPath = Path.Combine(Path.GetTempPath(), "mos_excel_variant_score.log");

        public static bool Miss(string reason)
        {
            ExcelScoreExplanation.Note(reason);
            Diag("Miss: " + (reason ?? string.Empty));
            return false;
        }

        public static void DiagReset()
        {
            try
            {
                File.WriteAllText(DiagLogPath, string.Empty);
            }
            catch
            {
            }
        }

        public static void Diag(string message)
        {
            string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
                + " " + (message ?? string.Empty);
            try
            {
                File.AppendAllText(DiagLogPath, line + Environment.NewLine);
            }
            catch
            {
            }

            Debug.WriteLine("[VariantScore] " + line);
        }

        public static string TryRead(Func<object> getter)
        {
            try
            {
                object value = getter();
                return value == null ? "<null>" : Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                return "ERR:" + ex.GetType().Name + ":" + ex.Message;
            }
        }

        public static Excel.Application GetApp()
        {
            try
            {
                return (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                return new Excel.Application { Visible = false };
            }
        }

        public static Excel.Workbook GetWorkbook(Excel.Application excelApp, string filePath)
        {
            if (excelApp?.Workbooks == null)
                return null;

            string fileName = string.IsNullOrEmpty(filePath) ? null : Path.GetFileName(filePath);
            foreach (Excel.Workbook workbook in excelApp.Workbooks)
            {
                try
                {
                    if (!string.IsNullOrEmpty(filePath) &&
                        string.Equals(workbook.FullName, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        return workbook;
                    }

                    if (!string.IsNullOrEmpty(fileName) &&
                        string.Equals(workbook.Name, fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        return workbook;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        public static Excel.Worksheet FindSheet(Excel.Workbook workbook, string sheetName)
        {
            if (workbook == null || string.IsNullOrEmpty(sheetName))
                return null;

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return sheet;
            }

            return null;
        }

        public static string NormalizeFormula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return string.Empty;

            string text = formula.Trim();
            if (text.StartsWith("=", StringComparison.Ordinal))
                text = text.Substring(1);
            text = text.Replace(" ", "");
            text = text.Replace("\"", "");
            text = text.Replace("'", "");
            return text.ToUpperInvariant();
        }

        public static string CellFormula(Excel.Range cell)
        {
            if (cell == null)
                return string.Empty;

            try
            {
                object formula = cell.HasFormula ? cell.Formula : null;
                return formula?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static string CellText(Excel.Range cell)
        {
            if (cell == null)
                return string.Empty;

            try
            {
                object text = cell.Text ?? cell.Value2;
                return Convert.ToString(text, CultureInfo.CurrentCulture)?.Trim() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static bool FormulaContainsAll(string formula, params string[] tokens)
        {
            string normalized = NormalizeFormula(formula);
            if (string.IsNullOrEmpty(normalized))
                return false;

            foreach (string token in tokens)
            {
                if (normalized.IndexOf(NormalizeFormula(token), StringComparison.Ordinal) < 0)
                    return false;
            }

            return true;
        }

        public static bool IsConsecutiveNumbers(Excel.Range range, double startValue)
        {
            if (range == null)
                return false;

            int rowCount = range.Rows.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                object value = range.Cells[i, 1].Value2;
                if (value == null || !double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out double number))
                    return false;
                if (Math.Abs(number - (startValue + i - 1)) > 0.001)
                    return false;
            }

            return true;
        }

        /// <summary>
        /// 先頭セルを起点に、末尾の数字だけが1ずつ増える連番（例: AF2606001, AF2606002）か見る。
        /// </summary>
        public static bool IsConsecutiveSerials(Excel.Range range)
        {
            if (range == null)
                return false;

            string first = CellText(range.Cells[1, 1]);
            if (!TrySplitSerial(first, out string prefix, out string digits))
                return false;

            if (!long.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out long startNumber))
                return false;

            int width = digits.Length;
            int rowCount = range.Rows.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                string expected = prefix + (startNumber + i - 1).ToString("D" + width, CultureInfo.InvariantCulture);
                string actual = CellText(range.Cells[i, 1]);
                if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            return true;
        }

        private static bool TrySplitSerial(string text, out string prefix, out string digits)
        {
            prefix = string.Empty;
            digits = string.Empty;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = text.Trim();
            int last = text.Length;
            while (last > 0 && char.IsDigit(text[last - 1]))
                last--;

            if (last == text.Length)
                return false;

            prefix = text.Substring(0, last);
            digits = text.Substring(last);
            return digits.Length > 0;
        }

        public static Excel.ListObject FirstTable(Excel.Worksheet sheet)
        {
            if (sheet?.ListObjects == null || sheet.ListObjects.Count == 0)
                return null;
            return sheet.ListObjects[1];
        }

        public static Excel.Chart FindChart(Excel.Worksheet sheet)
        {
            if (sheet == null)
                return null;

            if (sheet.ChartObjects().Count > 0)
            {
                Excel.ChartObject chartObject = sheet.ChartObjects(1);
                return chartObject.Chart;
            }

            return null;
        }

        public static string NormalizeAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return string.Empty;

            return Regex.Replace(address, @"\$", "").Replace(" ", "").ToUpperInvariant();
        }

        public static string TryGetBuiltinProperty(Excel.Workbook workbook, string propertyName)
        {
            try
            {
                dynamic properties = workbook.BuiltinDocumentProperties;
                dynamic property = properties[propertyName];
                return Convert.ToString(property.Value) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public static void Release(params object[] comObjects)
        {
            if (comObjects == null)
                return;

            foreach (object comObject in comObjects)
            {
                if (comObject == null)
                    continue;
                try
                {
                    Marshal.ReleaseComObject(comObject);
                }
                catch
                {
                }
            }
        }
    }
}
