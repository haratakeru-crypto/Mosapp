using System.Collections.Generic;

namespace MOS_Word_app
{
    /// <summary>
    /// 類題の総合点。最初に解けた20問は35点、それ以降の正解は20点。
    /// 未実装・不正解は0点。
    /// </summary>
    public static class WordPracticeScore
    {
        public const int HighValueCount = 20;
        public const int HighValuePoints = 35;
        public const int LowValuePoints = 20;
        public const int PassingScoreThreshold = 700;

        public static List<bool> FlattenResults(Dictionary<int, List<bool>> projectResults)
        {
            var flattened = new List<bool>();
            if (projectResults == null)
                projectResults = new Dictionary<int, List<bool>>();

            foreach (int projectId in VariantProblemWorkbook.GetProjectIds())
            {
                int taskCount = VariantProblemWorkbook.GetTaskCount(projectId);
                projectResults.TryGetValue(projectId, out List<bool> results);
                for (int i = 0; i < taskCount; i++)
                {
                    bool passed = results != null && i < results.Count && results[i];
                    flattened.Add(passed);
                }
            }

            return flattened;
        }

        public static int Compute(IEnumerable<bool> results)
        {
            int remainingHigh = HighValueCount;
            int total = 0;
            if (results == null)
                return 0;

            foreach (bool passed in results)
            {
                if (!passed)
                    continue;

                if (remainingHigh > 0)
                {
                    total += HighValuePoints;
                    remainingHigh--;
                }
                else
                {
                    total += LowValuePoints;
                }
            }

            return total;
        }
    }
}
