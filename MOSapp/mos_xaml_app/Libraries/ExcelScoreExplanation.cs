using System;
using System.Collections.Generic;
using System.Text;

namespace Libraries
{
    /// <summary>
    /// 採点の正誤は変えず、×の理由だけを学生向けの文にする。
    /// </summary>
    public static class ExcelScoreExplanation
    {
        public const string RequirementMissText = "求められている状態になっていないため、×になりました。";
        public const string UnavailableText = "この課題は確認できませんでした。";
        public const string UnknownKind = "unknown";
        public const string OutsideRangeKind = "outside-range";

        /// <summary>破壊的操作の動作名（未知）。一覧組み立て用。</summary>
        public const string UnknownStudentAction = "操作をした";

        /// <summary>許可範囲外編集の動作名。一覧組み立て用。</summary>
        public const string OutsideRangeStudentAction = "範囲のセルを編集した";

        /// <summary>破壊的操作の理由が取れなかったときの全文。</summary>
        public const string UnknownStudentText =
            "この課題の作業中に、この課題では使わない操作をした記録があるため、×になりました。";

        [ThreadStatic]
        private static string _checkerReason;

        /// <summary>タスク採点の直前に呼び、前の課題の理由を消す。</summary>
        public static void ClearCheckerReason()
        {
            _checkerReason = null;
        }

        /// <summary>チェッカーが false を返す直前に、学生向けの理由を1件足す。</summary>
        public static void Note(string reason)
        {
            if (string.IsNullOrWhiteSpace(reason))
                return;
            string line = reason.Trim();
            _checkerReason = string.IsNullOrEmpty(_checkerReason)
                ? line
                : _checkerReason + Environment.NewLine + line;
        }

        /// <summary>記録された理由を取り出し、次の課題に持ち越さない。</summary>
        public static string TakeCheckerReason()
        {
            string reason = _checkerReason;
            _checkerReason = null;
            return reason;
        }

        /// <summary>
        /// チェッカー結果に破壊的操作の判定を重ねる。戻り値の正誤は従来と同じです。
        /// </summary>
        public static bool Apply(int projectId, int taskId, bool checkerResult, out string reasonText)
        {
            return Apply(projectId, taskId, checkerResult, TakeCheckerReason(), out reasonText);
        }

        /// <summary>
        /// チェッカー結果に破壊的操作の判定を重ねる。戻り値の正誤は従来と同じです。
        /// 要件ミスと破壊的操作の両方があるときは、両方の理由を改行でつなぎます。
        /// </summary>
        public static bool Apply(int projectId, int taskId, bool checkerResult, string checkerReason, out string reasonText)
        {
            reasonText = null;
            if (projectId <= 0 || taskId <= 0)
            {
                if (!checkerResult)
                    reasonText = RequirementMissText;
                return checkerResult;
            }

            try
            {
                ExcelValidationExemptFlags exemptFlags = ExcelTaskValidationConfig.GetExemptFlags(projectId, taskId);
                int attemptNo = ExcelTaskAttemptRegistry.GetAttempt(projectId, taskId);
                bool violation = ExcelLogReader.TryCollectNonExemptViolations(
                    projectId,
                    taskId,
                    attemptNo,
                    exemptFlags,
                    out string firstLogMessage,
                    out List<string> studentActions);

                if (violation)
                {
                    if (checkerResult)
                    {
                        string line = $"P{projectId}-T{taskId} {firstLogMessage}";
                        System.Diagnostics.Debug.WriteLine($"[ExcelScoreExplanation] Destructive validation failed: {line}");
                        ExcelLogReader.AppendDestructiveError(projectId, taskId, attemptNo, line);
                    }

                    string destructiveText = FormatDestructiveStudentReasons(studentActions);
                    if (!checkerResult)
                    {
                        string checkerText = string.IsNullOrWhiteSpace(checkerReason)
                            ? RequirementMissText
                            : checkerReason.Trim();
                        reasonText = checkerText + Environment.NewLine + destructiveText;
                    }
                    else
                    {
                        reasonText = destructiveText;
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExcelScoreExplanation] Apply error: {ex.Message}");
                int attemptNo = ExcelTaskAttemptRegistry.GetAttempt(projectId, taskId);
                ExcelLogReader.AppendDestructiveError(projectId, taskId, attemptNo, $"P{projectId}-T{taskId} 例外: {ex.Message}");
                reasonText = UnavailableText;
                return false;
            }

            if (!checkerResult)
            {
                reasonText = string.IsNullOrWhiteSpace(checkerReason)
                    ? RequirementMissText
                    : checkerReason.Trim();
                return false;
            }

            return true;
        }

        /// <summary>
        /// 破壊的操作の動作名一覧を学生向け文にする。前置きは先頭1文のみ。
        /// </summary>
        public static string FormatDestructiveStudentReasons(IList<string> actions)
        {
            if (actions == null || actions.Count == 0)
                return UnknownStudentText;

            var cleaned = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string action in actions)
            {
                if (string.IsNullOrWhiteSpace(action))
                    continue;
                string line = action.Trim();
                if (seen.Add(line))
                    cleaned.Add(line);
            }

            if (cleaned.Count == 0)
                return UnknownStudentText;

            if (cleaned.Count == 1)
                return FormatRecordedAction(cleaned[0]);

            var sb = new StringBuilder();
            sb.Append(UnknownStudentText);
            foreach (string action in cleaned)
            {
                sb.Append(Environment.NewLine);
                sb.Append("・");
                sb.Append(action);
            }
            return sb.ToString();
        }

        /// <summary>
        /// 操作種別の種類キーと、学生向けの短い動作名（「〜した」）を返す。
        /// </summary>
        public static void DescribeOperation(ExcelOperationType op, out string kind, out string studentText)
        {
            switch (op)
            {
                case ExcelOperationType.EditCellValue:
                    kind = "EditCellValue";
                    studentText = "セルの内容を変えた";
                    return;
                case ExcelOperationType.EditCellFormula:
                    kind = "EditCellFormula";
                    studentText = "数式を変えた";
                    return;
                case ExcelOperationType.EditCellFormat:
                    kind = "EditCellFormat";
                    studentText = "セルの書式を変えた";
                    return;
                case ExcelOperationType.InsertRows:
                case ExcelOperationType.DeleteRows:
                    kind = "rows";
                    studentText = "行を挿入または削除した";
                    return;
                case ExcelOperationType.InsertColumns:
                case ExcelOperationType.DeleteColumns:
                    kind = "cols";
                    studentText = "列を挿入または削除した";
                    return;
                case ExcelOperationType.InsertShapeOrImage:
                case ExcelOperationType.DeleteShapeOrImage:
                    kind = "shape-count";
                    studentText = "図形や画像を追加または削除した";
                    return;
                case ExcelOperationType.MoveOrResizeShape:
                    kind = "MoveOrResizeShape";
                    studentText = "図形や画像の位置または大きさを変えた";
                    return;
                case ExcelOperationType.InsertHyperlink:
                    kind = "InsertHyperlink";
                    studentText = "ハイパーリンクを変えた";
                    return;
                case ExcelOperationType.SortOrFilter:
                    kind = "SortOrFilter";
                    studentText = "並べ替えやフィルターを変えた";
                    return;
                case ExcelOperationType.SetTableStyle:
                    kind = "SetTableStyle";
                    studentText = "テーブルのスタイルを変えた";
                    return;
                case ExcelOperationType.ResizeTable:
                    kind = "ResizeTable";
                    studentText = "テーブルの範囲を変えた";
                    return;
                case ExcelOperationType.AddConditionalFormat:
                    kind = "AddConditionalFormat";
                    studentText = "条件付き書式を変えた";
                    return;
                case ExcelOperationType.AddChart:
                    kind = "AddChart";
                    studentText = "グラフを追加した";
                    return;
                case ExcelOperationType.SetPrintArea:
                case ExcelOperationType.SetPrintTitle:
                case ExcelOperationType.SetHeaderFooter:
                case ExcelOperationType.SetPageBreak:
                case ExcelOperationType.SetPageOrientation:
                case ExcelOperationType.SetPageMargins:
                case ExcelOperationType.SetPageScaling:
                    kind = "print";
                    studentText = "印刷の設定を変えた";
                    return;
                case ExcelOperationType.SetFreezePanes:
                    kind = "SetFreezePanes";
                    studentText = "ウィンドウ枠の固定を変えた";
                    return;
                case ExcelOperationType.SetWorkbookProperty:
                    kind = "SetWorkbookProperty";
                    studentText = "文書の情報を変えた";
                    return;
                case ExcelOperationType.ManageNamedRange:
                    kind = "ManageNamedRange";
                    studentText = "名前の定義を変えた";
                    return;
                case ExcelOperationType.ImportExternalData:
                    kind = "ImportExternalData";
                    studentText = "外部データを取り込んだ";
                    return;
                default:
                    kind = UnknownKind;
                    studentText = UnknownStudentAction;
                    return;
            }
        }

        public static string FormatRecordedAction(string action)
        {
            return "この課題の作業中に、この課題では使わない" + action + "記録があるため、×になりました。";
        }
    }
}
