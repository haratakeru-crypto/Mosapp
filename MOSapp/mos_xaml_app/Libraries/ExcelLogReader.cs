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

        /// <summary>指定したプロジェクトに関連する破壊的操作エラーログのみをクリアします。</summary>
        public static void ClearDestructiveLogForProject(int projectId)
        {
            try
            {
                string path = GetDestructiveLogPath();
                if (!File.Exists(path)) return;

                string prefix = $"{projectId},";
                var lines = File.ReadAllLines(path);
                var keptLines = lines.Where(l => !l.StartsWith(prefix)).ToList();

                File.WriteAllLines(path, keptLines);
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

        /// <summary>指定したプロジェクトに関連する操作ログ（%TEMP%\mos_excel_log.txt）をクリアします。</summary>
        public static void ClearOperationLogForProject(int projectId)
        {
            try
            {
                string path = GetLogFilePath();
                if (!File.Exists(path)) return;

                var lines = File.ReadAllLines(path);
                var keptLines = new List<string>();
                int currentLineProjectId = -1;

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    // [TaskStart] 行のチェック
                    int taskStartIdx = line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase);
                    if (taskStartIdx >= 0)
                    {
                        ParseTaskStart(line, out currentLineProjectId, out _, out _);
                        if (currentLineProjectId != projectId)
                        {
                            keptLines.Add(line);
                        }
                        continue;
                    }

                    // 明示的なプロジェクト接頭辞（例: [Task 1-2-1]）のチェック
                    string projectPrefix = $"[Task {projectId}-";
                    if (line.IndexOf(projectPrefix, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // リセット対象プロジェクトの行なのでスキップ
                        continue;
                    }

                    // 文脈上のプロジェクトIDが一致しない（または不明な）場合は保持
                    if (currentLineProjectId != projectId)
                    {
                        keptLines.Add(line);
                    }
                }

                File.WriteAllLines(path, keptLines);
            }
            catch { }
        }

        public static void ClearOperationLog()
        {
            try
            {
                string path = GetLogFilePath();
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }

        public static string GetDiagnosticLogPath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_excel_addin_diag.txt");
        }

        /// <summary>VSTO アドイン自体の動作ログ（%TEMP%\mos_excel_addin_diag.txt）をクリアします。</summary>
        public static void ClearDiagnosticLog()
        {
            try
            {
                string path = GetDiagnosticLogPath();
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
            // 破壊的操作の判定は、常に「最新のタスク試行」のみを対象とする
            var ops = GetOperationsForTask(projectId, taskId, attemptNo, latestOnly: true);
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

        private static List<(string Type, string Detail)> GetOperationsForTask(int projectId, int taskId, int attemptNo, bool latestOnly = false)
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
                        
                        if (inTarget && latestOnly)
                        {
                            result.Clear();
                        }
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

        /// <summary>
        /// 指定されたグラフが作成された際の選択範囲（Selection）をログから取得する。
        /// </summary>
        public static string GetChartCreationSelection(int projectId, int taskId, int attemptNo, string chartName, string targetSheetName = null)
        {
            var ops = GetOperationsForTask(projectId, taskId, attemptNo);
            string finalSelection = null;

            string normChart = NormalizeChartName(chartName);
            // グラフ番号部分のみ抽出 (例: "5年間売上 グラフ 1" -> "グラフ1")
            string chartNumOnly = Regex.Match(normChart, @"グラフ\d+").Value;

            foreach (var op in ops)
            {
                if (!string.Equals(op.Type, "AddChart", StringComparison.OrdinalIgnoreCase)) continue;

                string detail = op.Detail;
                if (string.IsNullOrEmpty(detail)) continue;

                var nameMatch = Regex.Match(detail, @"Name=(.*?) Selection=");
                var selectionMatch = Regex.Match(detail, @"Selection=(.*)$");

                if (nameMatch.Success && selectionMatch.Success)
                {
                    string loggedName = nameMatch.Groups[1].Value.Trim();
                    string loggedSelection = selectionMatch.Groups[1].Value.Trim();

                    string normLogged = NormalizeChartName(loggedName);
                    
                    // シート名が指定されている場合は、Selection内のシート名を確認
                    if (!string.IsNullOrEmpty(targetSheetName))
                    {
                        // 全角半角無視してシート名が含まれているか確認
                        if (!NormalizeChartName(loggedSelection).Contains(NormalizeChartName(targetSheetName) + "!"))
                        {
                            continue;
                        }
                    }

                    // 名前の一致確認 (番号部分が一致するか、または全体が含まれているか)
                    bool nameMatches = normLogged.Contains(normChart) || normChart.Contains(normLogged);
                    if (!string.IsNullOrEmpty(chartNumOnly) && normLogged.Contains(chartNumOnly))
                    {
                        nameMatches = true;
                    }

                    if (nameMatches)
                    {
                        finalSelection = loggedSelection;
                    }
                }
            }
            return finalSelection;
        }

        private static string NormalizeChartName(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            
            char[] chars = input.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] >= '０' && chars[i] <= '９')
                    chars[i] = (char)(chars[i] - '０' + '0');
                else if (chars[i] >= 'Ａ' && chars[i] <= 'Ｚ')
                    chars[i] = (char)(chars[i] - 'Ａ' + 'A');
                else if (chars[i] >= 'ａ' && chars[i] <= 'ｚ')
                    chars[i] = (char)(chars[i] - 'ａ' + 'a');
            }
            string normalized = new string(chars).ToUpper().Replace("'", "");
            
            // 空白・制御文字削除
            return Regex.Replace(normalized, @"\s+", "");
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
