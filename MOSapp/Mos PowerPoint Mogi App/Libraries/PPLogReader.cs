using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Libraries
{
    /// <summary>
    /// PowerPoint VSTO アドインで生成されたログファイルを読み込むユーティリティ。
    /// メインログ（mos_ppt_log.txt）に加え、5-1/10-4/11-7 用の採点証跡（mos_ppt_task_evidence.txt）を扱う。
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
        /// 採点用証跡ログのパス（%TEMP%\mos_ppt_task_evidence.txt）。
        /// 5-1・10-4・11-7 など、単体プロジェクトリセット後も採点に必要な行だけを VSTO が追記する。
        /// <see cref="ClearLog"/> では消えない。全プロジェクトリセット時に <see cref="ClearTaskEvidence"/> で消す。
        /// </summary>
        public static string GetTaskEvidenceLogPath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_ppt_task_evidence.txt");
        }

        /// <summary>
        /// 現在タスク共有ファイルのパスを取得（%TEMP%\mos_ppt_current_task.txt）。
        /// 試験アプリが現在の ProjectId,TaskId を書き、VSTO アドインが読み取る。
        /// </summary>
        public static string GetCurrentTaskFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_ppt_current_task.txt");
        }

        /// <summary>破壊的操作ログのパスを取得（%TEMP%\mos_ppt_destructive_errors.log）</summary>
        public static string GetDestructiveLogPath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_ppt_destructive_errors.log");
        }

        /// <summary>スナップショットファイルのパスを取得（%TEMP%\mos_ppt_snapshot.txt）</summary>
        public static string GetSnapshotPath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_ppt_snapshot.txt");
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
        /// メイン操作ログ（mos_ppt_log.txt）のみクリアする。
        /// 採点用証跡は <see cref="ClearTaskEvidenceForProject"/> でプロジェクト単位に削除する。
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

        /// <summary>採点用証跡ログ（mos_ppt_task_evidence.txt）をクリアする。全プロジェクトリセット時に呼び出す。</summary>
        public static void ClearTaskEvidence()
        {
            try
            {
                string path = GetTaskEvidenceLogPath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] Error clearing task evidence: " + ex.Message);
            }
        }

        /// <summary>
        /// 単体プロジェクトリセット時、そのプロジェクトのログ依存採点タスクに対応する証跡行だけを削除する。
        /// 5→5-1 印刷、10→10-4 グレースケール、11→11-7 印刷。他プロジェクトでは何もしない。
        /// </summary>
        public static void ClearTaskEvidenceForProject(int projectId)
        {
            string[] markers = GetEvidenceLineMarkersToRemoveForProject(projectId);
            if (markers == null || markers.Length == 0)
                return;

            string path = GetTaskEvidenceLogPath();
            if (!File.Exists(path))
                return;

            try
            {
                var lines = File.ReadAllLines(path, Encoding.UTF8);
                bool LineMatchesAnyMarker(string line)
                {
                    if (string.IsNullOrEmpty(line)) return false;
                    foreach (string m in markers)
                    {
                        if (line.IndexOf(m, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                    }
                    return false;
                }

                var kept = lines.Where(l => !LineMatchesAnyMarker(l)).ToArray();
                if (kept.Length == lines.Length)
                    return;

                if (kept.Length == 0)
                {
                    File.Delete(path);
                    return;
                }

                File.WriteAllLines(path, kept, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] ClearTaskEvidenceForProject: " + ex.Message);
            }
        }

        /// <summary>証跡ファイルから削除する行に含まれるべきマーカー（Logger が書く形式と一致）。</summary>
        private static string[] GetEvidenceLineMarkersToRemoveForProject(int projectId)
        {
            switch (projectId)
            {
                case 5:
                    return new[] { "[Task5-1] Print" };
                case 10:
                    return new[] { "[Task10-4] Grayscale" };
                case 11:
                    return new[] { "[Task11-7] Print" };
                default:
                    return null;
            }
        }

        /// <summary>破壊的操作ログをクリアする。</summary>
        public static void ClearDestructiveLog()
        {
            try
            {
                string path = GetDestructiveLogPath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] Error clearing destructive log: " + ex.Message);
            }
        }

        /// <summary>スナップショットファイルをクリアする。</summary>
        public static void ClearSnapshot()
        {
            try
            {
                string path = GetSnapshotPath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] Error clearing snapshot: " + ex.Message);
            }
        }

        /// <summary>
        /// 証跡ログまたはメインログに [Task10-4] Grayscale が含まれるか（移行前のセッションはメインログのみの可能性あり）。
        /// </summary>
        public static bool HasTask10_4GrayscaleExecuted()
        {
            return HasGradingEvidenceMarker("[Task10-4] Grayscale");
        }

        /// <summary>証跡またはメインログに 5-1 の印刷記録（[Task5-1] Print）が含まれるか。</summary>
        public static bool HasTask5_1PrintExecuted()
        {
            return HasGradingEvidenceMarker("[Task5-1] Print");
        }

        /// <summary>証跡またはメインログに 11-7 の印刷記録（[Task11-7] Print）が含まれるか。</summary>
        public static bool HasTask11_7PrintExecuted()
        {
            return HasGradingEvidenceMarker("[Task11-7] Print");
        }

        /// <summary>採点用証跡を優先し、無ければ従来の mos_ppt_log.txt を検索する。</summary>
        private static bool HasGradingEvidenceMarker(string marker)
        {
            return FileContainsMarker(GetTaskEvidenceLogPath(), marker)
                || FileContainsMarker(GetLogFilePath(), marker);
        }

        private static bool FileContainsMarker(string path, string marker)
        {
            if (string.IsNullOrEmpty(marker) || string.IsNullOrEmpty(path) || !File.Exists(path))
                return false;
            try
            {
                return File.ReadAllLines(path).Any(line =>
                    line != null && line.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] FileContainsMarker: " + ex.Message);
                return false;
            }
        }

        /// <summary>ログに 8-4 オーディオ設定記録（[Task8-4] Audio）が含まれるか。</summary>
        public static bool HasTask8_4AudioExecuted()
        {
            return HasLogLineContaining("[Task8-4] Audio");
        }

        /// <summary>
        /// 指定タスクの区間（[TaskStart] project-task から次の [TaskStart] 直前まで）に、指定マーカーが含まれるか。
        /// 例: Task8-4 の Audio は、過去ログが残っていると誤判定しやすいため区間内検索を推奨。
        /// </summary>
        public static bool HasMarkerWithinTask(int projectId, int taskId, string marker)
        {
            if (string.IsNullOrEmpty(marker))
                return false;
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return false;
            try
            {
                string[] lines = File.ReadAllLines(path);
                bool inTarget = false;
                int curP = -1, curT = -1;
                foreach (string line in lines)
                {
                    if (line == null) continue;
                    if (line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ParseTaskStart(line, out curP, out curT);
                        inTarget = (curP == projectId && curT == taskId);
                        continue;
                    }
                    if (!inTarget) continue;
                    if (line.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] HasMarkerWithinTask: " + ex.Message);
            }
            return false;
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

        /// <summary>ログに 7-4 Kiosk 設定記録（[Task7-4] Kiosk）が含まれるか。</summary>
        public static bool HasTask7_4KioskExecuted()
        {
            return HasLogLineContaining("[Task7-4] Kiosk");
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
