using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Libraries
{
    public enum ExcelBulkResetScope
    {
        Textbook = 0,
        Variant = 1,
        All = 2
    }

    public sealed class ExcelResetTarget
    {
        public int GroupId { get; set; }
        public int ProjectId { get; set; }
        public int VariantSetNo { get; set; }
        public string Label { get; set; }
        public string WorkingPath { get; set; }
        public string TemplatePath { get; set; }
        public string InitialPath { get; set; }
        public bool IsVariant => VariantSetNo > 0;
    }

    public sealed class ExcelBulkResetResult
    {
        public int TextbookSuccess { get; set; }
        public int TextbookFail { get; set; }
        public int VariantSuccess { get; set; }
        public int VariantFail { get; set; }
        public List<string> FailedProjects { get; } = new List<string>();
        public ExcelResetTarget ReopenTarget { get; set; }

        public int TotalSuccess => TextbookSuccess + VariantSuccess;
        public int TotalFail => TextbookFail + VariantFail;
    }

    /// <summary>
    /// 演習 / 類題 / すべての作業ファイルを、単体リセットをループせず一括で戻す。
    /// COM のクローズと再オープンは呼び出し側（UI/STA）が1回ずつ行う。
    /// </summary>
    public static class ExcelBulkResetCoordinator
    {
        public const int DefaultVariantSetNo = 1;

        public static ExcelBulkResetResult CopyTargets(
            IList<ExcelResetTarget> targets,
            ExcelResetTarget currentOpenTarget,
            Action<string> reportProgress)
        {
            var result = new ExcelBulkResetResult();
            if (targets == null)
                return result;

            foreach (ExcelResetTarget target in targets)
            {
                try
                {
                    reportProgress?.Invoke(target.Label + " をリセット中");
                    ExcelProjectResetHelper.RestoreFromTemplate(
                        target.TemplatePath,
                        target.WorkingPath,
                        target.IsVariant ? null : target.InitialPath);
                    ExcelProjectResetHelper.ClearLogsAfterSuccessfulCopy(target.ProjectId, target.VariantSetNo);
                    if (target.IsVariant)
                        result.VariantSuccess++;
                    else
                        result.TextbookSuccess++;

                    if (SameTarget(target, currentOpenTarget))
                        result.ReopenTarget = target;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[ExcelBulkResetCoordinator] " + target.Label + ": " + ex.Message);
                    result.FailedProjects.Add(target.Label);
                    if (target.IsVariant)
                        result.VariantFail++;
                    else
                        result.TextbookFail++;
                }
            }

            return result;
        }

        public static bool SameTarget(ExcelResetTarget left, ExcelResetTarget right)
        {
            if (left == null || right == null)
                return false;
            return left.GroupId == right.GroupId
                && left.ProjectId == right.ProjectId
                && left.VariantSetNo == right.VariantSetNo;
        }
    }
}
