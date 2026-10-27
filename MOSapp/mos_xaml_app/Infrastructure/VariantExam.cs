using System;
using System.Collections.Generic;
using System.IO;
using MOSExcelMogiApp.Models;
using Newtonsoft.Json.Linq;

namespace MOSExcelMogiApp.Infrastructure
{
    /// <summary>類題採点で使うパス・ライブラリ名・結果の並びを一箇所にまとめる。</summary>
    public static class VariantExam
    {
        public static ExamResultScope Scope(int variantSetNo) => ExamResultScope.Variant(variantSetNo);

        public static string GetLibraryName(JObject config, int groupId, int projectId, int variantSetNo)
        {
            JToken entry = config?["practiceVariants"]?[groupId.ToString()]?[projectId.ToString()]?[variantSetNo.ToString()];
            string libraryName = entry?["library"]?.ToString();
            if (!string.IsNullOrWhiteSpace(libraryName))
                return libraryName;

            return $"ExcelChecker{groupId}_{projectId}_PV{variantSetNo}";
        }

        public static string GetWorkingFilePath(JObject config, int groupId, int projectId, int variantSetNo)
        {
            JToken entry = config?["practiceVariants"]?[groupId.ToString()]?[projectId.ToString()]?[variantSetNo.ToString()];
            string configured = entry?["excelFile"]?.ToString();
            return DataPathHelper.ResolveVariantWorkingFilePath(groupId, projectId, variantSetNo, configured);
        }

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

        public static bool TryValidateFiles(
            JObject config,
            int groupId,
            int variantSetNo,
            out List<string> missingFiles)
        {
            missingFiles = new List<string>();
            foreach (int projectId in VariantProblemWorkbook.GetProjectIds())
            {
                string path = GetWorkingFilePath(config, groupId, projectId, variantSetNo);
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    missingFiles.Add($"Project {projectId}");
            }

            return missingFiles.Count == 0;
        }
    }
}
