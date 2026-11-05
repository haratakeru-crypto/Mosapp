using System;
using System.Collections.Generic;

namespace MOS_PowerPoint_app
{
    /// <summary>類題の解答済み／あとで見直す。同時には片方だけ立つ。</summary>
    public static class PowerPointTaskReviewState
    {
        public const int LostPointsPerCompletedWrong = 40;

        public static bool ShowLostPoints(bool isVariantMode) => isVariantMode;

        public static Dictionary<int, bool[]> Clone(Dictionary<int, bool[]> source)
        {
            var clone = new Dictionary<int, bool[]>();
            if (source == null)
                return clone;

            foreach (var kv in source)
                clone[kv.Key] = kv.Value != null ? (bool[])kv.Value.Clone() : Array.Empty<bool>();
            return clone;
        }

        public static bool RequiresDispositionBeforeScoring(bool isVariantMode) => isVariantMode;

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
                if (GetFlag(completed, task.ProjectId, task.TaskId)
                    || GetFlag(flagged, task.ProjectId, task.TaskId))
                    continue;

                unmarked.Add(new UnmarkedTask
                {
                    ProjectId = task.ProjectId,
                    TaskId = task.TaskId
                });
            }

            return unmarked;
        }

        public static string FormatUnmarkedMessage(IList<UnmarkedTask> unmarked)
        {
            if (unmarked == null || unmarked.Count == 0)
                return string.Empty;

            return $"解答済み／あとで見直すが未選択の問題が {unmarked.Count} 問あります。{Environment.NewLine}このまま採点しますか？";
        }

        public static int CountCompletedWrong(
            Dictionary<int, List<bool>> results,
            Dictionary<int, bool[]> completed)
        {
            if (results == null || completed == null)
                return 0;

            int count = 0;
            foreach (var kv in results)
            {
                if (kv.Value == null)
                    continue;

                for (int i = 0; i < kv.Value.Count; i++)
                {
                    int taskId = i + 1;
                    if (kv.Value[i])
                        continue;
                    if (GetFlag(completed, kv.Key, taskId))
                        count++;
                }
            }

            return count;
        }

        public static string FormatLostPointsMessage(int completedWrongCount)
        {
            if (completedWrongCount <= 0)
                return string.Empty;

            int points = completedWrongCount * LostPointsPerCompletedWrong;
            return $"解答済みにした問題のうち{completedWrongCount}問が不正解でした。約{points}点分の損となります";
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

        public static void ToggleCompleted(
            Dictionary<int, bool[]> completed,
            Dictionary<int, bool[]> flagged,
            int projectId,
            int taskId,
            int minLength)
        {
            SetExclusive(
                completed,
                flagged,
                projectId,
                taskId,
                GetFlag(completed, projectId, taskId) ? Disposition.None : Disposition.Completed,
                minLength);
        }

        public static void ToggleReviewLater(
            Dictionary<int, bool[]> completed,
            Dictionary<int, bool[]> flagged,
            int projectId,
            int taskId,
            int minLength)
        {
            SetExclusive(
                completed,
                flagged,
                projectId,
                taskId,
                GetFlag(flagged, projectId, taskId) ? Disposition.None : Disposition.ReviewLater,
                minLength);
        }

        private enum Disposition
        {
            None,
            Completed,
            ReviewLater
        }

        private static void SetExclusive(
            Dictionary<int, bool[]> completed,
            Dictionary<int, bool[]> flagged,
            int projectId,
            int taskId,
            Disposition disposition,
            int minLength)
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

            completedArray[index] = disposition == Disposition.Completed;
            flaggedArray[index] = disposition == Disposition.ReviewLater;
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
