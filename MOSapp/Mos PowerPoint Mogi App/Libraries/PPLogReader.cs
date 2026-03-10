using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Libraries
{
    /// <summary>
    /// PowerPoint VSTO アドインで生成されたログファイルを読み込むユーティリティ。
    /// 採点で 10-4 グレースケール等のログを参照する。
    /// </summary>
    public static class PPLogReader
    {
        /// <summary>
        /// ログファイルのパスを取得（%TEMP%\mos_ppt_log.txt）
        /// </summary>
        public static string GetLogFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_ppt_log.txt");
        }

        /// <summary>
        /// 現在タスク共有ファイルのパスを取得（%TEMP%\mos_ppt_current_task.txt）。
        /// 試験アプリが現在の ProjectId,TaskId を書き、VSTO アドインが読み取る。
        /// </summary>
        public static string GetCurrentTaskFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_ppt_current_task.txt");
        }

        /// <summary>
        /// 現在タスク共有ファイルを削除する。リセット時に呼び出す。
        /// </summary>
        public static void ClearCurrentTaskFile()
        {
            try
            {
                string path = GetCurrentTaskFilePath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] Error clearing current task file: " + ex.Message);
            }
        }

        /// <summary>
        /// VSTO アドインのログファイルをクリアする。リセット時に呼び出す。
        /// </summary>
        public static void ClearLog()
        {
            try
            {
                string path = GetLogFilePath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] Error clearing log: " + ex.Message);
            }
        }

        /// <summary>
        /// ログに [Task10-4] Grayscale が 1 行でも含まれるか（10-4 グレースケール操作の記録あり）
        /// </summary>
        public static bool HasTask10_4GrayscaleExecuted()
        {
            return HasLogLineContaining("[Task10-4] Grayscale");
        }

        /// <summary>ログに 5-1 の印刷記録（[Task5-1] Print）が含まれるか。</summary>
        public static bool HasTask5_1PrintExecuted()
        {
            return HasLogLineContaining("[Task5-1] Print");
        }

        /// <summary>ログに 11-7 の印刷記録（[Task11-7] Print）が含まれるか。</summary>
        public static bool HasTask11_7PrintExecuted()
        {
            return HasLogLineContaining("[Task11-7] Print");
        }

        /// <summary>ログに 8-4 オーディオ設定記録（[Task8-4] Audio）が含まれるか。</summary>
        public static bool HasTask8_4AudioExecuted()
        {
            return HasLogLineContaining("[Task8-4] Audio");
        }

        /// <summary>ログに 7-2 スライド再利用記録（[Task7-2] ReuseSlides）が含まれるか。</summary>
        public static bool HasTask7_2ReuseSlidesExecuted()
        {
            return HasLogLineContaining("[Task7-2] ReuseSlides");
        }

        /// <summary>ログに 7-3 アウトラインから挿入記録（[Task7-3] InsertFromOutline）が含まれるか。</summary>
        public static bool HasTask7_3InsertFromOutlineExecuted()
        {
            return HasLogLineContaining("[Task7-3] InsertFromOutline");
        }

        /// <summary>ログに 10-1 ドキュメント検査記録（[Task10-1] DocumentInspector）が含まれるか。</summary>
        public static bool HasTask10_1DocumentInspectorExecuted()
        {
            return HasLogLineContaining("[Task10-1] DocumentInspector");
        }

        /// <summary>ログに 10-7 レイアウト複製記録（[Task10-7] LayoutDuplicate）が含まれるか。</summary>
        public static bool HasTask10_7LayoutDuplicateExecuted()
        {
            return HasLogLineContaining("[Task10-7] LayoutDuplicate");
        }

        private static bool HasLogLineContaining(string marker)
        {
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return false;
            try
            {
                return File.ReadAllLines(path).Any(line =>
                    line != null && line.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] Error reading log: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// ログファイルからすべての行を読み込む（デバッグ・表示用）
        /// </summary>
        public static List<string> ReadAllLines()
        {
            var result = new List<string>();
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return result;
            try
            {
                result.AddRange(File.ReadAllLines(path));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] Error reading log: " + ex.Message);
            }
            return result;
        }

        /// <summary>
        /// 指定タスク区間内の操作行を取得する。[TaskStart] で区切った区間のうち、指定 projectId-taskId の区間内の [Op] 行の内容（[Op] 以降）を返す。
        /// </summary>
        public static List<string> GetOperationsForTask(int projectId, int taskId)
        {
            var result = new List<string>();
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return result;
            try
            {
                string[] lines = File.ReadAllLines(path);
                int currentProject = -1, currentTask = -1;
                foreach (string line in lines)
                {
                    if (line == null) continue;
                    if (line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ParseTaskStart(line, out currentProject, out currentTask);
                        continue;
                    }
                    if (currentProject != projectId || currentTask != taskId)
                        continue;
                    int opIdx = line.IndexOf("[Op]", StringComparison.OrdinalIgnoreCase);
                    if (opIdx < 0) continue;
                    string afterOp = line.Substring(opIdx + 4).Trim();
                    if (afterOp.Length > 0)
                        result.Add(afterOp);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] GetOperationsForTask: " + ex.Message);
            }
            return result;
        }

        /// <summary>
        /// 指定タスク区間に、許可リストに含まれない操作が 1 件でもあれば true。許可リストは操作タイプ（RibbonCommand, ShapePositionChange 等）の集合。
        /// ログが無い・空の場合は false（厳格判定しない）。
        /// </summary>
        public static bool HasDisallowedOperations(int projectId, int taskId, HashSet<string> allowedOperationTypes)
        {
            if (allowedOperationTypes == null)
                return false;
            var ops = GetOperationsForTask(projectId, taskId);
            foreach (string opLine in ops)
            {
                string type = GetOperationType(opLine);
                if (string.IsNullOrEmpty(type)) continue;
                if (!allowedOperationTypes.Contains(type))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// 指定タスク区間に ShapePositionChange が 1 件でもあれば true。
        /// </summary>
        public static bool HasShapePositionChange(int projectId, int taskId)
        {
            var ops = GetOperationsForTask(projectId, taskId);
            foreach (string opLine in ops)
            {
                if (opLine.IndexOf("ShapePositionChange", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        private static void ParseTaskStart(string line, out int projectId, out int taskId)
        {
            projectId = -1;
            taskId = -1;
            int startIdx = line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase);
            if (startIdx < 0) return;
            string part = line.Substring(startIdx + 11).Trim();
            var tokens = part.Split(new[] { '-', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length >= 2 && int.TryParse(tokens[0].Trim(), out projectId) && int.TryParse(tokens[1].Trim(), out taskId))
                return;
            projectId = -1;
            taskId = -1;
        }


        private static string GetOperationType(string opLine)
        {
            if (string.IsNullOrWhiteSpace(opLine)) return "";
            int space = opLine.IndexOf(' ');
            return space > 0 ? opLine.Substring(0, space).Trim() : opLine.Trim();
        }
    }
}
