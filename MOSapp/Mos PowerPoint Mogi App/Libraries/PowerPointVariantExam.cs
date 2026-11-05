using System.Collections.Generic;

namespace MOS_PowerPoint_app
{
    /// <summary>類題採点結果の並び・スコープを一箇所にまとめる。</summary>
    public static class PowerPointVariantExam
    {
        public static List<bool> FlattenResults(Dictionary<int, List<bool>> projectResults)
        {
            var flattened = new List<bool>();
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

        public static bool HasAllProjects(Dictionary<int, List<bool>> projectResults)
        {
            var expected = VariantProblemWorkbook.GetProjectIds();
            if (expected == null || expected.Count == 0 || projectResults == null)
                return false;

            foreach (int projectId in expected)
            {
                if (!projectResults.ContainsKey(projectId))
                    return false;
            }

            return true;
        }
    }
}
