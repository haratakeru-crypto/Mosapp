using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Newtonsoft.Json;

namespace MOSExcelMogiApp.Guide
{
    /// <summary>
    /// References/GuideSteps 配下の manifest / PNG を安全に読み込む。
    /// </summary>
    public static class GuideStepsRepository
    {
        private static readonly Regex SafeFileNameRegex = new Regex(
            @"^[A-Za-z0-9_\-\.]+\.(png|jpg|jpeg|gif|bmp)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public static string GetGuideStepsRoot()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "GuideSteps");
        }

        public static string GetTaskFolder(int groupId, int projectId, int taskId)
        {
            return Path.Combine(
                GetGuideStepsRoot(),
                $"Group{groupId}",
                $"Project{projectId}",
                $"Task{taskId}");
        }

        /// <summary>
        /// 指定プロジェクトで有効なタスク手順だけを返す（壊れたタスクはスキップ）。
        /// </summary>
        public static IReadOnlyList<AnswerGuideTask> LoadProjectTasks(int groupId, int projectId)
        {
            var result = new List<AnswerGuideTask>();
            string projectFolder = Path.Combine(
                GetGuideStepsRoot(),
                $"Group{groupId}",
                $"Project{projectId}");

            if (!Directory.Exists(projectFolder))
            {
                Debug.WriteLine($"[GuideStepsRepository] Project folder missing: {projectFolder}");
                return result;
            }

            foreach (string taskDir in Directory.EnumerateDirectories(projectFolder, "Task*")
                         .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    string folderName = Path.GetFileName(taskDir);
                    if (folderName == null ||
                        !folderName.StartsWith("Task", StringComparison.OrdinalIgnoreCase) ||
                        !int.TryParse(folderName.Substring(4), out int folderTaskId) ||
                        folderTaskId <= 0)
                    {
                        Debug.WriteLine($"[GuideStepsRepository] Skip invalid task folder: {taskDir}");
                        continue;
                    }

                    AnswerGuideTask task = TryLoadTask(groupId, projectId, folderTaskId, taskDir);
                    if (task != null)
                        result.Add(task);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[GuideStepsRepository] Skip task folder due to error: {taskDir} ({ex.Message})");
                }
            }

            return result.OrderBy(t => t.TaskId).ToList();
        }

        public static AnswerGuideTask TryLoadTask(int groupId, int projectId, int taskId)
        {
            return TryLoadTask(groupId, projectId, taskId, GetTaskFolder(groupId, projectId, taskId));
        }

        private static AnswerGuideTask TryLoadTask(int groupId, int projectId, int taskId, string taskDir)
        {
            if (!Directory.Exists(taskDir))
                return null;

            string manifestPath = Path.Combine(taskDir, "manifest.json");
            if (!File.Exists(manifestPath))
            {
                Debug.WriteLine($"[GuideStepsRepository] manifest missing: {manifestPath}");
                return null;
            }

            AnswerGuideManifest manifest;
            try
            {
                string json = File.ReadAllText(manifestPath);
                manifest = JsonConvert.DeserializeObject<AnswerGuideManifest>(json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GuideStepsRepository] manifest parse failed: {manifestPath} ({ex.Message})");
                return null;
            }

            if (manifest == null)
            {
                Debug.WriteLine($"[GuideStepsRepository] manifest null: {manifestPath}");
                return null;
            }

            if (manifest.GroupId != groupId ||
                manifest.ProjectId != projectId ||
                manifest.TaskId != taskId)
            {
                Debug.WriteLine(
                    $"[GuideStepsRepository] manifest id mismatch: expected {groupId}-{projectId}-{taskId}, " +
                    $"got {manifest.GroupId}-{manifest.ProjectId}-{manifest.TaskId} ({manifestPath})");
                return null;
            }

            if (manifest.Steps == null || manifest.Steps.Count == 0)
            {
                Debug.WriteLine($"[GuideStepsRepository] no steps: {manifestPath}");
                return null;
            }

            var loaded = new AnswerGuideTask { TaskId = taskId };
            for (int i = 0; i < manifest.Steps.Count; i++)
            {
                AnswerGuideStep step = manifest.Steps[i];
                if (step == null)
                {
                    Debug.WriteLine($"[GuideStepsRepository] null step at {i}: {manifestPath}");
                    return null;
                }

                if (!IsSafeRelativeImageFileName(step.Image))
                {
                    Debug.WriteLine($"[GuideStepsRepository] unsafe image name '{step.Image}' at {i}: {manifestPath}");
                    return null;
                }

                string imagePath = Path.GetFullPath(Path.Combine(taskDir, step.Image));
                string taskDirFull = Path.GetFullPath(taskDir)
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    + Path.DirectorySeparatorChar;
                if (!imagePath.StartsWith(taskDirFull, StringComparison.OrdinalIgnoreCase) ||
                    !File.Exists(imagePath))
                {
                    Debug.WriteLine($"[GuideStepsRepository] image missing or escaped: {imagePath}");
                    return null;
                }

                loaded.Steps.Add(new AnswerGuideLoadedStep
                {
                    ImagePath = imagePath,
                    Caption = step.Caption?.Trim() ?? string.Empty
                });
            }

            return loaded;
        }

        private static bool IsSafeRelativeImageFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return false;
            if (fileName.IndexOfAny(new[] { '/', '\\' }) >= 0)
                return false;
            if (fileName.Contains(".."))
                return false;
            if (Path.IsPathRooted(fileName))
                return false;
            return SafeFileNameRegex.IsMatch(fileName);
        }
    }
}
