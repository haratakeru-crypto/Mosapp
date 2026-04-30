using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Libraries
{
    /// <summary>
    /// Excel VSTO アドインのログ読み取り・破壊的操作判定ユーティリティ。
    /// </summary>
    public static class ExcelLogReader
    {
        private static readonly Regex CellRegex = new Regex(@"^([A-Z]+)(\d+)$", RegexOptions.Compiled);

        public static string GetLogFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_excel_log.txt");
        }

        public static string GetCurrentTaskFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_excel_current_task.txt");
        }

        public static void ClearCurrentTaskFile()
        {
            try
            {
                string path = GetCurrentTaskFilePath();
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }

        /// <summary>破壊的操作エラーログ（採点時の不合格理由）。PowerPoint の mos_ppt_destructive_errors.log に対応。</summary>
        public static string GetDestructiveLogPath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_excel_destructive_errors.log");
        }

        public static void AppendDestructiveError(int projectId, int taskId, int attemptNo, string message)
        {
            try
            {
                string line = $"{projectId},{taskId},{attemptNo}:{message}{Environment.NewLine}";
                File.AppendAllText(GetDestructiveLogPath(), line);
            }
            catch { }
        }

        public static void ClearDestructiveLog()
        {
            try
            {
                string path = GetDestructiveLogPath();
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }

        /// <summary>VSTO が追記する操作ログ（%TEMP%\mos_excel_log.txt）。リセット時に破壊的操作判定の入力をクリアする。</summary>
        public static void ClearOperationLog()
        {
            try
            {
                string path = GetLogFilePath();
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }


        /// <summary>
        /// 方式A（プロジェクト1想定）: 免除カテゴリに含まれない操作が1件でもあれば違反。
        /// 許可範囲が定義されているタスクでは、範囲編集は許可範囲内のみ通過（それ以外は許可範囲外として違反）。
        /// </summary>
        public static bool TryGetFirstNonExemptViolation(
            int projectId,
            int taskId,
            int attemptNo,
            ExcelValidationExemptFlags exemptFlags,
            out string message)
        {
            message = null;
            var ops = GetOperationsForTask(projectId, taskId, attemptNo);
            bool useRangeGate = ExcelTaskValidationConfig.ShouldDenyOutsideAllowedRanges(projectId, taskId);
            List<string> allowedRanges = useRangeGate
                ? ExcelTaskValidationConfig.GetAllowedRanges(projectId, taskId)
                : null;

            foreach (var op in ops)
            {
                if (!Enum.TryParse(op.Type, out ExcelOperationType opType))
                {
                    message = $"未対応の操作種別: {op.Type} (detail: {op.Detail})";
                    return true;
                }

                if (ExcelTaskValidationConfig.IsOperationExempt(opType, exemptFlags))
                    continue;

                if (useRangeGate && allowedRanges != null && allowedRanges.Count > 0 && IsRangeEditType(op.Type))
                {
                    if (!TryParseSheetAndAddress(op.Detail, out string sheetName, out string address))
                    {
                        message = $"許可範囲外の編集: {op.Type}（アドレス解釈不可: {op.Detail}）";
                        return true;
                    }

                    if (IsAddressWithinAllowedRanges(sheetName, address, allowedRanges))
                        continue;

                    message = $"許可範囲外の編集: {op.Type} {op.Detail}";
                    return true;
                }

                message = $"免除外の操作: {opType} (detail: {op.Detail})";
                return true;
            }

            return false;
        }


        private static bool IsRangeEditType(string operationType)
        {
            return string.Equals(operationType, "EditCellValue", StringComparison.OrdinalIgnoreCase)
                || string.Equals(operationType, "EditCellFormula", StringComparison.OrdinalIgnoreCase)
                || string.Equals(operationType, "EditCellFormat", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryParseSheetAndAddress(string detail, out string sheetName, out string address)
        {
            sheetName = "";
            address = "";
            if (string.IsNullOrWhiteSpace(detail)) return false;

            int sep = detail.IndexOf('!');
            if (sep <= 0 || sep >= detail.Length - 1) return false;
            sheetName = detail.Substring(0, sep).Trim();
            address = detail.Substring(sep + 1).Trim().Replace("$", "");
            return !string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(address);
        }

        private static bool IsAddressWithinAllowedRanges(string sheetName, string address, List<string> allowedRanges)
        {
            var areas = SplitAddressAreas(address);
            foreach (var area in areas)
            {
                bool areaAllowed = false;
                foreach (string allowed in allowedRanges)
                {
                    if (!TryParseAllowedRange(allowed, out string allowedSheet, out string allowedAddress))
                        continue;
                    if (!string.Equals(allowedSheet, sheetName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (IsAreaWithin(allowedAddress, area))
                    {
                        areaAllowed = true;
                        break;
                    }
                }
                if (!areaAllowed) return false;
            }
            return true;
        }

        private static List<string> SplitAddressAreas(string address)
        {
            return address
                .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrEmpty(x))
                .ToList();
        }

        private static bool TryParseAllowedRange(string value, out string sheet, out string address)
        {
            sheet = "";
            address = "";
            int sep = value.IndexOf('!');
            if (sep <= 0 || sep >= value.Length - 1) return false;
            sheet = value.Substring(0, sep).Trim();
            address = value.Substring(sep + 1).Trim().Replace("$", "");
            return !string.IsNullOrEmpty(sheet) && !string.IsNullOrEmpty(address);
        }

        private static bool IsAreaWithin(string allowedAddress, string targetArea)
        {
            if (!TryGetRangeBounds(allowedAddress, out int aCol1, out int aRow1, out int aCol2, out int aRow2))
                return false;
            if (!TryGetRangeBounds(targetArea, out int tCol1, out int tRow1, out int tCol2, out int tRow2))
                return false;

            return tCol1 >= aCol1 && tCol2 <= aCol2 && tRow1 >= aRow1 && tRow2 <= aRow2;
        }

        private static bool TryGetRangeBounds(string area, out int col1, out int row1, out int col2, out int row2)
        {
            col1 = row1 = col2 = row2 = 0;
            string[] parts = area.Split(':');
            if (parts.Length == 1)
            {
                if (!TryParseCell(parts[0], out col1, out row1)) return false;
                col2 = col1;
                row2 = row1;
                return true;
            }
            if (parts.Length == 2)
            {
                if (!TryParseCell(parts[0], out col1, out row1)) return false;
                if (!TryParseCell(parts[1], out col2, out row2)) return false;
                if (col1 > col2) (col1, col2) = (col2, col1);
                if (row1 > row2) (row1, row2) = (row2, row1);
                return true;
            }
            return false;
        }

        private static bool TryParseCell(string value, out int col, out int row)
        {
            col = 0;
            row = 0;
            var m = CellRegex.Match(value.Trim().ToUpperInvariant());
            if (!m.Success) return false;
            col = ColumnNameToNumber(m.Groups[1].Value);
            row = int.Parse(m.Groups[2].Value);
            return col > 0 && row > 0;
        }

        private static int ColumnNameToNumber(string name)
        {
            int n = 0;
            foreach (char c in name)
            {
                n = (n * 26) + (c - 'A' + 1);
            }
            return n;
        }

        private static List<(string Type, string Detail)> GetOperationsForTask(int projectId, int taskId, int attemptNo)
        {
            var result = new List<(string Type, string Detail)>();
            string path = GetLogFilePath();
            if (!File.Exists(path)) return result;

            try
            {
                string[] lines = File.ReadAllLines(path);
                bool inTarget = false;
                int curP = -1, curT = -1, curA = 1;
                string explicitPrefix = $"[Task {projectId}-{taskId}-{attemptNo}]";

                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    if (line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ParseTaskStart(line, out curP, out curT, out curA);
                        inTarget = (curP == projectId && curT == taskId && curA == attemptNo);
                        continue;
                    }

                    bool hasExplicitTaskPrefix = line.IndexOf(explicitPrefix, StringComparison.OrdinalIgnoreCase) >= 0;
                    if (!inTarget && !hasExplicitTaskPrefix) continue;

                    int opIdx = line.IndexOf("[Op]", StringComparison.OrdinalIgnoreCase);
                    if (opIdx < 0) continue;
                    string payload = line.Substring(opIdx + 4).Trim();
                    if (string.IsNullOrEmpty(payload)) continue;

                    int firstSpace = payload.IndexOf(' ');
                    string opType = firstSpace > 0 ? payload.Substring(0, firstSpace).Trim() : payload.Trim();
                    string detail = firstSpace > 0 ? payload.Substring(firstSpace + 1).Trim() : "";
                    result.Add((opType, detail));
                }
            }
            catch { }

            return result;
        }

        private static void ParseTaskStart(string line, out int projectId, out int taskId, out int attemptNo)
        {
            projectId = -1;
            taskId = -1;
            attemptNo = 1;
            int idx = line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return;

            string part = line.Substring(idx + 11).Trim();
            var tokens = part.Split(new[] { '-', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length < 2) return;

            if (!int.TryParse(tokens[0].Trim(), out projectId)) { projectId = -1; return; }
            if (!int.TryParse(tokens[1].Trim(), out taskId)) { taskId = -1; return; }
            if (tokens.Length >= 3)
            {
                int.TryParse(tokens[2].Trim(), out attemptNo);
                if (attemptNo < 1) attemptNo = 1;
            }
        }
    }
}
