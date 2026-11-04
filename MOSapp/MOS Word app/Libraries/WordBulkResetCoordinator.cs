using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Libraries
{
    public enum WordBulkResetScope
    {
        Textbook = 0,
        Variant = 1,
        All = 2
    }

    public sealed class WordResetTarget
    {
        public int GroupId { get; set; }
        public int ProjectId { get; set; }
        public int VariantSetNo { get; set; }
        public string Label { get; set; }
        public string WorkingPath { get; set; }
        public string TemplatePath { get; set; }
        public bool IsVariant => VariantSetNo > 0;
    }

    public sealed class WordBulkResetResult
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
    /// 演習 / 類題 / すべての作業ファイルを一括で戻す。Word のクローズは呼び出し側が行う。
    /// </summary>
    public static class WordBulkResetCoordinator
    {
        public const int DefaultVariantSetNo = 1;

        public static WordBulkResetResult CopyTargets(
            IList<WordResetTarget> targets,
            Action<string> reportProgress)
        {
            var result = new WordBulkResetResult();
            if (targets == null || targets.Count == 0)
                return result;

            foreach (WordResetTarget target in targets)
            {
                try
                {
                    reportProgress?.Invoke(target.Label + " をリセット中");
                    if (target.IsVariant)
                    {
                        MOS_Word_app.WordProjectResetHelper.RestoreVariantProjectFiles(
                            target.GroupId,
                            target.ProjectId,
                            target.VariantSetNo);
                    }
                    else
                    {
                        MOS_Word_app.WordProjectResetHelper.RestoreTextbookProjectFiles(
                            target.GroupId,
                            target.ProjectId);
                    }

                    MOS_Word_app.WordProjectResetHelper.ClearLogsAfterSuccessfulCopy(
                        target.ProjectId,
                        target.VariantSetNo);

                    if (target.IsVariant)
                        result.VariantSuccess++;
                    else
                        result.TextbookSuccess++;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("[WordBulkResetCoordinator] " + target.Label + ": " + ex.Message);
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
