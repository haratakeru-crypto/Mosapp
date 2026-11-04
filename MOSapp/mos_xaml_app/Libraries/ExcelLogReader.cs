using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Libraries
{
    /// <summary>
    /// Excel VSTO アドインのログ読み取り・破壊的操作判定ユーティリティ。
    /// </summary>
    public static class ExcelLogReader
    {
        private static readonly Regex CellRegex = new Regex(@"^([A-Z]+)(\d+)$", RegexOptions.Compiled);
        private static string[] _logLinesCache;

        /// <summary>採点時にログ読み取りを高速化するため、全行をメモリにキャッシュします。</summary>
        public static void SetLogLinesCache(string[] lines)
        {
            _logLinesCache = lines;
        }

        /// <summary>キャッシュをクリアします。</summary>
        public static void ClearLogLinesCache()
        {
            _logLinesCache = null;
        }

        public static string GetLogFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_excel_log.txt");
        }

        public static string GetCurrentTaskFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_excel_current_task.txt");
        }

        public static string GetBoundaryFlushRequestPath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_excel_flush_boundary.txt");
        }

        public static string GetLogSchemaMarkerPath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_excel_log_schema_v2.marker");
        }

        /// <summary>
        /// 教材と類題のログキーを分離したスキーマへ一度だけ移行する。
        /// 旧ログは教材/類題を判別できないため、初回のみ操作ログと破壊ログを破棄する。
        /// </summary>
        public static void EnsureLogSchemaV2()
        {
            try
            {
                string marker = GetLogSchemaMarkerPath();
                if (File.Exists(marker))
                    return;

                ClearOperationLog();
                ClearDestructiveLog();
                ClearCurrentTaskFile();
                CancelPendingBoundaryFlush();
                ClearLogLinesCache();
                File.WriteAllText(marker, DateTime.UtcNow.ToString("o"), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ExcelLogReader] EnsureLogSchemaV2: " + ex.Message);
            }
        }

        /// <summary>
        /// その場採点の開始時、いま開いているタスクの未記録差分だけをアドインへ書き出させる。
        /// 採点対象プロジェクトと現在タスクが違うときは何もしない。
        /// </summary>
        public static void RequestOpenTaskBoundaryFlush(int scoringProjectId, int timeoutMs = 1500, int variantSetNo = -1)
        {
            if (!TryReadCurrentTask(out TaskLogContext context) || !context.IsValid)
                return;
            if (context.ProjectId != scoringProjectId)
                return;
            if (variantSetNo >= 0 && context.VariantSetNo != variantSetNo)
                return;

            TryRequestBoundaryFlush(context, timeoutMs, out _);
        }

        /// <summary>
        /// レビュー遷移前に、現在タスクの未記録差分をVSTOへ確定させる。
        /// 要求ファイルが削除されたときだけ成功とする。現在タスクが無い、または時間内に消えないときは false。
        /// </summary>
        public static bool TryRequestCurrentTaskBoundaryFlush(int timeoutMs, out string detail)
        {
            if (!TryReadCurrentTask(out TaskLogContext context) || !context.IsValid)
            {
                detail = "no-current-task";
                return false;
            }

            return TryRequestBoundaryFlush(context, timeoutMs, out detail);
        }

        private static bool TryRequestBoundaryFlush(TaskLogContext context, int timeoutMs, out string detail)
        {
            detail = context.FormatTaskToken();
            string path = GetBoundaryFlushRequestPath();
            try
            {
                File.WriteAllText(path, context.FormatCurrentTaskLine(), Encoding.UTF8);
                var sw = Stopwatch.StartNew();
                while (sw.ElapsedMilliseconds < timeoutMs)
                {
                    if (!File.Exists(path))
                    {
                        detail += " flushed";
                        return true;
                    }
                    Thread.Sleep(50);
                }

                detail += " timeout";
                return false;
            }
            catch (Exception ex)
            {
                detail += " error";
                Debug.WriteLine("[ExcelLogReader] TryRequestBoundaryFlush: " + ex.Message);
                return false;
            }
        }

        public static void CancelPendingBoundaryFlush()
        {
            try
            {
                string path = GetBoundaryFlushRequestPath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch { }
        }

        public static void CancelPendingBoundaryFlushIfMatches(int projectId, int variantSetNo)
        {
            try
            {
                string path = GetBoundaryFlushRequestPath();
                if (!File.Exists(path))
                    return;
                string line = File.ReadAllText(path).Trim();
                if (!TaskLogFormat.TryParseCurrentTaskLine(line, out TaskLogContext context)
                    || context.MatchesProjectScope(projectId, variantSetNo))
                {
                    File.Delete(path);
                }
            }
            catch { }
        }

        private static bool TryReadCurrentTask(out TaskLogContext context)
        {
            context = default(TaskLogContext);
            try
            {
                string path = GetCurrentTaskFilePath();
                if (!File.Exists(path))
                    return false;
                string line = File.ReadAllText(path).Trim();
                return TaskLogFormat.TryParseCurrentTaskLine(line, out context) && context.IsValid;
            }
            catch
            {
                return false;
            }
        }

        public static void ClearCurrentTaskFileIfMatches(int projectId, int variantSetNo)
        {
            if (TryReadCurrentTask(out TaskLogContext context) && context.MatchesProjectScope(projectId, variantSetNo))
                ClearCurrentTaskFile();
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

        public static void AppendDestructiveError(int projectId, int taskId, int attemptNo, string message, int variantSetNo = 0)
        {
            if (!ExcelDiagFlags.IsVerbose)
                return;
            try
            {
                var context = TaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo);
                string line = context.FormatDestructivePrefix() + message + Environment.NewLine;
                File.AppendAllText(GetDestructiveLogPath(), line);
            }
            catch { }
        }

        /// <summary>指定したプロジェクトに関連する破壊的操作エラーログのみをクリアします。</summary>
        public static void ClearDestructiveLogForProject(int projectId)
        {
            ClearDestructiveLogForProject(projectId, 0);
        }

        public static void ClearDestructiveLogForProject(int projectId, int variantSetNo)
        {
            try
            {
                string path = GetDestructiveLogPath();
                if (!File.Exists(path)) return;

                var lines = File.ReadAllLines(path);
                var keptLines = new List<string>();
                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;
                    if (TaskLogFormat.TryParseDestructivePrefix(line, out TaskLogContext context)
                        && context.MatchesProjectScope(projectId, variantSetNo))
                        continue;
                    keptLines.Add(line);
                }

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
            ClearOperationLogForProject(projectId, 0);
        }

        public static void ClearOperationLogForProject(int projectId, int variantSetNo)
        {
            try
            {
                string path = GetLogFilePath();
                if (!File.Exists(path)) return;

                var lines = File.ReadAllLines(path);
                var keptLines = new List<string>();
                TaskLogContext currentLineContext = default(TaskLogContext);
                bool currentLineValid = false;

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    if (line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        currentLineValid = TaskLogFormat.TryParseTaskStartLine(line, out currentLineContext);
                        if (!(currentLineValid && currentLineContext.MatchesProjectScope(projectId, variantSetNo)))
                            keptLines.Add(line);
                        continue;
                    }

                    if (TaskLogFormat.TryParseExplicitTaskPrefix(line, out TaskLogContext explicitContext)
                        && explicitContext.MatchesProjectScope(projectId, variantSetNo))
                    {
                        continue;
                    }

                    if (currentLineValid && currentLineContext.MatchesProjectScope(projectId, variantSetNo))
                        continue;

                    keptLines.Add(line);
                }

                File.WriteAllLines(path, keptLines);
                ClearLogLinesCache();
            }
            catch { }
        }

        public static void ClearLogsForProjectScope(int projectId, int variantSetNo)
        {
            ClearOperationLogForProject(projectId, variantSetNo);
            ClearDestructiveLogForProject(projectId, variantSetNo);
            ExcelTaskAttemptRegistry.ClearProject(projectId, variantSetNo);
            ClearLogLinesCache();
            ClearCurrentTaskFileIfMatches(projectId, variantSetNo);
            CancelPendingBoundaryFlushIfMatches(projectId, variantSetNo);
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

        /// <summary>
        /// VSTO 診断ログ。詳細は Debug、または %TEMP%\mos_excel_diag.on があるときだけ書く。
        /// 起動待ちの Startup completed / Baseline completed は常に残す。
        /// </summary>
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
            ExcelVstoReadiness.ResetAfterDiagnosticLogCleared();
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
            out string message,
            int variantSetNo = 0)
        {
            if (!TryCollectNonExemptViolations(projectId, taskId, attemptNo, exemptFlags, out string firstLogMessage, out _, variantSetNo))
            {
                message = null;
                return false;
            }

            message = firstLogMessage;
            return true;
        }

        /// <summary>
        /// 免除外の操作を、学生向けの理由文として種類ごとに集める。正誤は先頭の1件と同じです。
        /// </summary>
        public static bool TryCollectNonExemptViolations(
            int projectId,
            int taskId,
            int attemptNo,
            ExcelValidationExemptFlags exemptFlags,
            out string firstLogMessage,
            out List<string> studentReasons,
            int variantSetNo = 0)
        {
            firstLogMessage = null;
            studentReasons = new List<string>();
            // 破壊的操作の判定は、同一タスク内の全試行（セッション）の操作を合算して判定する。
            // 以前は latestOnly: true だったが、タスクを切り替えて戻るだけで違反が消えてしまうのを防ぐため false に変更。
            var ops = GetOperationsForTask(projectId, taskId, attemptNo, latestOnly: false, variantSetNo: variantSetNo);
            bool useRangeGate = ExcelTaskValidationConfig.ShouldDenyOutsideAllowedRanges(projectId, taskId);
            List<string> allowedRanges = useRangeGate
                ? ExcelTaskValidationConfig.GetAllowedRanges(projectId, taskId)
                : null;
            var seenKinds = new HashSet<string>(StringComparer.Ordinal);

            foreach (var op in ops)
            {
                if (!Enum.TryParse(op.Type, out ExcelOperationType opType))
                {
                    RememberViolation(
                        ref firstLogMessage,
                        studentReasons,
                        seenKinds,
                        $"未対応の操作種別: {op.Type} (detail: {op.Detail})",
                        ExcelScoreExplanation.UnknownKind,
                        ExcelScoreExplanation.UnknownStudentAction);
                    continue;
                }

                // タスク切替などで別シートに付いたレイアウト系ログは、対象シート外なら違反にしない
                if (ExcelTaskValidationConfig.ShouldIgnoreOffTargetSheetOperation(projectId, taskId, op.Detail))
                    continue;

                if (ExcelTaskValidationConfig.IsOperationExempt(opType, exemptFlags))
                    continue;

                if (IsSpuriousUsedRangeDimensionChangeForCellFormatTask(opType, op.Detail, exemptFlags))
                    continue;

                if (useRangeGate && allowedRanges != null && allowedRanges.Count > 0 && IsRangeEditType(op.Type))
                {
                    if (!TryParseSheetAndAddress(op.Detail, out string sheetName, out string address))
                    {
                        RememberViolation(
                            ref firstLogMessage,
                            studentReasons,
                            seenKinds,
                            $"許可範囲外の編集: {op.Type}（アドレス解釈不可: {op.Detail}）",
                            ExcelScoreExplanation.OutsideRangeKind,
                            ExcelScoreExplanation.OutsideRangeStudentAction);
                        continue;
                    }

                    if (IsAddressWithinAllowedRanges(sheetName, address, allowedRanges))
                        continue;

                    RememberViolation(
                        ref firstLogMessage,
                        studentReasons,
                        seenKinds,
                        $"許可範囲外の編集: {op.Type} {op.Detail}",
                        ExcelScoreExplanation.OutsideRangeKind,
                        ExcelScoreExplanation.OutsideRangeStudentAction);
                    continue;
                }

                ExcelScoreExplanation.DescribeOperation(opType, out string kind, out string studentText);
                RememberViolation(
                    ref firstLogMessage,
                    studentReasons,
                    seenKinds,
                    $"免除外の操作: {opType} (detail: {op.Detail})",
                    kind,
                    studentText);
            }

            return firstLogMessage != null;
        }

        private static void RememberViolation(
            ref string firstLogMessage,
            List<string> studentReasons,
            HashSet<string> seenKinds,
            string logMessage,
            string kind,
            string studentText)
        {
            if (firstLogMessage == null)
                firstLogMessage = logMessage;
            if (string.IsNullOrEmpty(kind) || !seenKinds.Add(kind))
                return;
            if (!string.IsNullOrEmpty(studentText))
                studentReasons.Add(studentText);
        }


        private static bool IsRangeEditType(string operationType)
        {
            return string.Equals(operationType, "EditCellValue", StringComparison.OrdinalIgnoreCase)
                || string.Equals(operationType, "EditCellFormula", StringComparison.OrdinalIgnoreCase)
                || string.Equals(operationType, "EditCellFormat", StringComparison.OrdinalIgnoreCase);
        }

        private static readonly Regex UsedRowCountDriftRegex = new Regex(@"Rows:(\d+)->(\d+)", RegexOptions.Compiled);
        private static readonly Regex UsedColCountDriftRegex = new Regex(@"Cols:(\d+)->(\d+)", RegexOptions.Compiled);

        /// <summary>
        /// 書式のみタスクで UsedRange 行数/列数の微小変化が InsertRows 等として記録された場合の誤検知を無視する。
        /// </summary>
        private static bool IsSpuriousUsedRangeDimensionChangeForCellFormatTask(
            ExcelOperationType opType,
            string detail,
            ExcelValidationExemptFlags exemptFlags)
        {
            if (!exemptFlags.HasFlag(ExcelValidationExemptFlags.CellFormatOnly))
                return false;

            if (opType != ExcelOperationType.InsertRows
                && opType != ExcelOperationType.DeleteRows
                && opType != ExcelOperationType.InsertColumns
                && opType != ExcelOperationType.DeleteColumns)
            {
                return false;
            }

            if (string.IsNullOrEmpty(detail))
                return false;

            Match rowMatch = UsedRowCountDriftRegex.Match(detail);
            if (rowMatch.Success
                && int.TryParse(rowMatch.Groups[1].Value, out int fromRow)
                && int.TryParse(rowMatch.Groups[2].Value, out int toRow)
                && Math.Abs(toRow - fromRow) <= 1)
            {
                return true;
            }

            Match colMatch = UsedColCountDriftRegex.Match(detail);
            if (colMatch.Success
                && int.TryParse(colMatch.Groups[1].Value, out int fromCol)
                && int.TryParse(colMatch.Groups[2].Value, out int toCol)
                && Math.Abs(toCol - fromCol) <= 1)
            {
                return true;
            }

            return false;
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

        private static List<(string Type, string Detail)> GetOperationsForTask(
            int projectId,
            int taskId,
            int attemptNo,
            bool latestOnly = false,
            int variantSetNo = 0)
        {
            var result = new List<(string Type, string Detail)>();
            string path = GetLogFilePath();
            if (!File.Exists(path)) return result;

            try
            {
                string[] lines = _logLinesCache;
                if (lines == null)
                {
                    if (!File.Exists(path)) return result;
                    lines = File.ReadAllLines(path);
                }

                bool inTarget = false;
                var target = TaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo);
                string explicitPrefix = target.FormatExplicitTaskPrefix();

                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    if (line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        inTarget = TaskLogFormat.TryParseTaskStartLine(line, out TaskLogContext startContext)
                            && startContext.ProjectId == projectId
                            && startContext.TaskId == taskId
                            && startContext.AttemptNo == attemptNo
                            && startContext.VariantSetNo == variantSetNo;

                        if (inTarget && latestOnly)
                        {
                            result.Clear();
                        }
                        continue;
                    }

                    bool hasExplicitTaskPrefix = false;
                    if (TaskLogFormat.TryParseExplicitTaskPrefix(line, out TaskLogContext explicitContext))
                    {
                        hasExplicitTaskPrefix = explicitContext.Equals(target);
                    }
                    else
                    {
                        hasExplicitTaskPrefix = variantSetNo == 0
                            && line.IndexOf(explicitPrefix, StringComparison.OrdinalIgnoreCase) >= 0;
                    }
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
        public static string GetChartCreationSelection(int projectId, int taskId, int attemptNo, string chartName, string targetSheetName = null, int variantSetNo = 0)
        {
            var ops = GetOperationsForTask(projectId, taskId, attemptNo, variantSetNo: variantSetNo);
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

    }

    /// <summary>
    /// タスク単位の再挑戦番号。未登録は 1。結果画面から×のタスクに戻ったときだけ増やす。
    /// 類題は V{set}-{project}-{task}、教材は {project}-{task}。
    /// </summary>
    public static class ExcelTaskAttemptRegistry
    {
        private static readonly object Sync = new object();
        private static readonly Dictionary<string, int> Attempts = new Dictionary<string, int>(StringComparer.Ordinal);

        private static string Key(int projectId, int taskId, int variantSetNo)
        {
            if (variantSetNo > 0)
                return "V" + variantSetNo + "-" + projectId + "-" + taskId;
            return projectId + "-" + taskId;
        }

        public static int GetAttempt(int projectId, int taskId)
        {
            return GetAttempt(projectId, taskId, 0);
        }

        public static int GetAttempt(int projectId, int taskId, int variantSetNo)
        {
            lock (Sync)
            {
                if (Attempts.TryGetValue(Key(projectId, taskId, variantSetNo), out int v) && v >= 1)
                    return v;
                return 1;
            }
        }

        public static void SetAttempt(int projectId, int taskId, int attemptNo)
        {
            SetAttempt(projectId, taskId, attemptNo, 0);
        }

        public static void SetAttempt(int projectId, int taskId, int attemptNo, int variantSetNo)
        {
            if (attemptNo < 1) attemptNo = 1;
            lock (Sync)
            {
                Attempts[Key(projectId, taskId, variantSetNo)] = attemptNo;
            }
        }

        public static int Increment(int projectId, int taskId)
        {
            return Increment(projectId, taskId, 0);
        }

        public static int Increment(int projectId, int taskId, int variantSetNo)
        {
            lock (Sync)
            {
                string key = Key(projectId, taskId, variantSetNo);
                int current = 1;
                if (Attempts.TryGetValue(key, out int v) && v >= 1)
                    current = v;
                int next = current + 1;
                Attempts[key] = next;
                return next;
            }
        }

        public static void ClearProject(int projectId)
        {
            ClearProject(projectId, 0);
        }

        public static void ClearProject(int projectId, int variantSetNo)
        {
            lock (Sync)
            {
                var keys = new List<string>();
                if (variantSetNo > 0)
                {
                    string prefix = "V" + variantSetNo + "-" + projectId + "-";
                    foreach (var key in Attempts.Keys)
                    {
                        if (key.StartsWith(prefix, StringComparison.Ordinal))
                            keys.Add(key);
                    }
                }
                else
                {
                    string prefix = projectId + "-";
                    foreach (var key in Attempts.Keys)
                    {
                        if (key.StartsWith("V", StringComparison.Ordinal))
                            continue;
                        if (key.StartsWith(prefix, StringComparison.Ordinal))
                            keys.Add(key);
                    }
                }
                foreach (var key in keys)
                    Attempts.Remove(key);
            }
        }
    }
}
