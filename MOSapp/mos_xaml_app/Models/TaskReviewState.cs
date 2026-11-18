using System;
using System.Collections.Generic;

namespace MOSExcelMogiApp.Models
{
    public enum TaskReviewDisposition
    {
        None = 0,
        Completed = 1,
        ReviewLater = 2
    }

    /// <summary>
    /// 解答済み／あとで見直すの排他状態。既存のプロジェクト別 bool 配列を共有する。
    /// </summary>
    public static class TaskReviewState
    {
        public const int LostPointsPerCompletedWrong = 40;

        public static bool RequiresDispositionBeforeScoring(bool isVariantMode, int variantSetNo)
        {
            return isVariantMode && variantSetNo == 1;
        }

        /// <summary>教材・類題とも結果画面に解答済み／あとで見直す列を出す。点数・優先バッジは別ゲート。</summary>
        public static bool ShowDispositionColumns(bool isVariantMode)
        {
            return true;
        }

        public static bool ShowLostPoints(bool isVariantMode)
        {
            return isVariantMode;
        }

        public static Dictionary<int, bool[]> Clone(Dictionary<int, bool[]> source)
        {
            var clone = new Dictionary<int, bool[]>();
            if (source == null)
                return clone;

            foreach (var kv in source)
                clone[kv.Key] = kv.Value != null ? (bool[])kv.Value.Clone() : Array.Empty<bool>();
            return clone;
        }

        public static bool GetFlag(Dictionary<int, bool[]> states, int projectId, int taskId)
        {
            if (states == null || taskId <= 0)
                return false;
            if (!states.TryGetValue(projectId, out bool[] array) || array == null)
                return false;

            int index = taskId - 1;
            return index >= 0 && index < array.Length && array[index];
        }

        public static TaskReviewDisposition Get(
            Dictionary<int, bool[]> completed,
            Dictionary<int, bool[]> flagged,
            int projectId,
            int taskId)
        {
            if (GetFlag(completed, projectId, taskId))
                return TaskReviewDisposition.Completed;
            if (GetFlag(flagged, projectId, taskId))
                return TaskReviewDisposition.ReviewLater;
            return TaskReviewDisposition.None;
        }

        public static void ToggleCompleted(
            Dictionary<int, bool[]> completed,
            Dictionary<int, bool[]> flagged,
            int projectId,
            int taskId,
            int minLength = 8)
        {
            bool currentlyCompleted = GetFlag(completed, projectId, taskId);
            SetExclusive(
                completed,
                flagged,
                projectId,
                taskId,
                currentlyCompleted ? TaskReviewDisposition.None : TaskReviewDisposition.Completed,
                minLength);
        }

        public static void ToggleReviewLater(
            Dictionary<int, bool[]> completed,
            Dictionary<int, bool[]> flagged,
            int projectId,
            int taskId,
            int minLength = 8)
        {
            bool currentlyFlagged = GetFlag(flagged, projectId, taskId);
            SetExclusive(
                completed,
                flagged,
                projectId,
                taskId,
                currentlyFlagged ? TaskReviewDisposition.None : TaskReviewDisposition.ReviewLater,
                minLength);
        }

        public static void SetExclusive(
            Dictionary<int, bool[]> completed,
            Dictionary<int, bool[]> flagged,
            int projectId,
            int taskId,
            TaskReviewDisposition disposition,
            int minLength = 8)
        {
            if (completed == null || flagged == null || projectId <= 0 || taskId <= 0)
                return;

            int length = Math.Max(Math.Max(minLength, 8), taskId);
            EnsureCapacity(completed, projectId, length);
            EnsureCapacity(flagged, projectId, length);

            bool[] completedArray = completed[projectId];
            bool[] flaggedArray = flagged[projectId];
            int index = taskId - 1;
            if (index < 0 || index >= completedArray.Length || index >= flaggedArray.Length)
                return;

            switch (disposition)
            {
                case TaskReviewDisposition.Completed:
                    completedArray[index] = true;
                    flaggedArray[index] = false;
                    break;
                case TaskReviewDisposition.ReviewLater:
                    completedArray[index] = false;
                    flaggedArray[index] = true;
                    break;
                default:
                    completedArray[index] = false;
                    flaggedArray[index] = false;
                    break;
            }
        }

        public sealed class UnmarkedTask
        {
            public int ProjectId { get; set; }
            public int TaskId { get; set; }
        }

        public static List<UnmarkedTask> FindUnmarked(
            IEnumerable<(int ProjectId, int TaskId)> tasks,
            Dictionary<int, bool[]> completed,
            Dictionary<int, bool[]> flagged)
        {
            var unmarked = new List<UnmarkedTask>();
            if (tasks == null)
                return unmarked;

            foreach (var task in tasks)
            {
                if (Get(completed, flagged, task.ProjectId, task.TaskId) == TaskReviewDisposition.None)
                {
                    unmarked.Add(new UnmarkedTask
                    {
                        ProjectId = task.ProjectId,
                        TaskId = task.TaskId
                    });
                }
            }

            return unmarked;
        }

        public static int CountCompletedWrong(
            Dictionary<int, List<bool>> results,
            Dictionary<int, bool[]> completed)
        {
            CountDispositions(results, completed, null, out _, out int completedWrongCount, out _, out _);
            return completedWrongCount;
        }

        /// <summary>
        /// 結果に含まれるタスクについて、解答済み／あとで見直すと、それぞれの間違えた問題数を数える。
        /// </summary>
        public static void CountDispositions(
            Dictionary<int, List<bool>> results,
            Dictionary<int, bool[]> completed,
            Dictionary<int, bool[]> flagged,
            out int completedCount,
            out int completedWrongCount,
            out int reviewLaterCount,
            out int reviewLaterWrongCount)
        {
            completedCount = 0;
            completedWrongCount = 0;
            reviewLaterCount = 0;
            reviewLaterWrongCount = 0;
            if (results == null)
                return;

            foreach (var kv in results)
            {
                if (kv.Value == null)
                    continue;

                for (int i = 0; i < kv.Value.Count; i++)
                {
                    int taskId = i + 1;
                    bool isWrong = !kv.Value[i];
                    bool isCompleted = GetFlag(completed, kv.Key, taskId);
                    bool isFlagged = GetFlag(flagged, kv.Key, taskId);
                    if (isCompleted)
                    {
                        completedCount++;
                        if (isWrong)
                            completedWrongCount++;
                    }
                    if (isFlagged)
                    {
                        reviewLaterCount++;
                        if (isWrong)
                            reviewLaterWrongCount++;
                    }
                }
            }
        }

        public static string FormatDispositionSummary(
            int completedCount,
            int completedWrongCount,
            int reviewLaterCount,
            int reviewLaterWrongCount)
        {
            return $"解答済み：{completedCount}問（うち間違えた問題：{completedWrongCount}問）　あとで見直す：{reviewLaterCount}問（うち間違えた問題：{reviewLaterWrongCount}問）";
        }

        public static string FormatLostPointsMessage(int completedWrongCount)
        {
            if (completedWrongCount <= 0)
                return string.Empty;

            int points = completedWrongCount * LostPointsPerCompletedWrong;
            return $"解答済みにした問題のうち{completedWrongCount}問が不正解でした。約{points}点分の損となります";
        }

        public static string FormatUnmarkedMessage(IList<UnmarkedTask> unmarked)
        {
            if (unmarked == null || unmarked.Count == 0)
                return string.Empty;

            return $"解答済み／あとで見直すが未選択の問題が {unmarked.Count} 問あります。{Environment.NewLine}このまま採点しますか？";
        }

        private static void EnsureCapacity(Dictionary<int, bool[]> states, int projectId, int length)
        {
            if (states == null || projectId <= 0 || length <= 0)
                return;

            if (!states.TryGetValue(projectId, out bool[] array) || array == null)
            {
                states[projectId] = new bool[length];
                return;
            }

            if (array.Length >= length)
                return;

            var resized = new bool[length];
            Array.Copy(array, resized, array.Length);
            states[projectId] = resized;
        }
    }
}
