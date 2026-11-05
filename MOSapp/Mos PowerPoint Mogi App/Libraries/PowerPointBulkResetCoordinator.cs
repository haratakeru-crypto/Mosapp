using System;
using System.Collections.Generic;
using System.Diagnostics;
using MOS_PowerPoint_app;

namespace Libraries
{
    public enum PowerPointBulkResetScope
    {
        Textbook = 0,
        Variant = 1,
        All = 2
    }

    public sealed class PowerPointResetTarget
    {
        public int GroupId { get; set; }
        public int ProjectId { get; set; }
        public int VariantSetNo { get; set; }
        public string Label { get; set; }
        public bool IsVariant => VariantSetNo > 0;
    }

    public sealed class PowerPointBulkResetResult
    {
        public int TextbookSuccess { get; set; }
        public int TextbookFail { get; set; }
        public int VariantSuccess { get; set; }
        public int VariantFail { get; set; }
        public List<string> FailedProjects { get; } = new List<string>();

        public int TotalSuccess => TextbookSuccess + VariantSuccess;
        public int TotalFail => TextbookFail + VariantFail;
    }

    /// <summary>
    /// 演習 / 類題 / すべての作業ファイルを一括で戻す。
    /// </summary>
    public static class PowerPointBulkResetCoordinator
    {
        public const int DefaultVariantSetNo = 1;

        public static PowerPointBulkResetResult CopyTargets(
            IList<PowerPointResetTarget> targets,
            Action<string> reportProgress)
        {
            var result = new PowerPointBulkResetResult();
            if (targets == null || targets.Count == 0)
                return result;

            foreach (PowerPointResetTarget target in targets)
            {
                try
                {
                    reportProgress?.Invoke(target.Label + " をリセット中");
                    if (target.IsVariant)
                    {
                        PowerPointProjectResetHelper.ResetVariantProject(
                            target.GroupId,
                            target.ProjectId,
                            target.VariantSetNo);
                        result.VariantSuccess++;
                    }
                    else
                    {
                        PowerPointProjectResetHelper.ResetProject(target.GroupId, target.ProjectId);
                        result.TextbookSuccess++;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[PowerPointBulkResetCoordinator] " + target.Label + ": " + ex.Message);
                    result.FailedProjects.Add(target.Label);
                    if (target.IsVariant)
                        result.VariantFail++;
                    else
                        result.TextbookFail++;
                }
            }

            return result;
        }
    }
}
