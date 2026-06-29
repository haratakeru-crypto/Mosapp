using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace Libraries
{
    /// <summary>
    /// PowerPoint VSTO アドインで生成されたログファイルを読み込むユーティリティ。
    /// メインログ（mos_ppt_log.txt）に加え、1-2/1-3/1-4/1-8/4-3/5-1/10-4/11-7 用の採点証跡（mos_ppt_task_evidence.txt）を扱う。
    /// </summary>
    public static class PPLogReader
    {
        private static readonly AsyncLocal<int?> _gradingProjectId = new AsyncLocal<int?>();
        private static readonly AsyncLocal<int?> _gradingTaskId = new AsyncLocal<int?>();
        private static readonly AsyncLocal<int?> _gradingAttemptNo = new AsyncLocal<int?>();

        public static void SetGradingContext(int projectId, int taskId, int attemptNo)
        {
            _gradingProjectId.Value = projectId;
            _gradingTaskId.Value = taskId;
            _gradingAttemptNo.Value = attemptNo;
        }

        public static void ClearGradingContext()
        {
            _gradingProjectId.Value = null;
            _gradingTaskId.Value = null;
            _gradingAttemptNo.Value = null;
        }

        /// <summary>
        /// ログファイルのパスを取得（%TEMP%\mos_ppt_log.txt）
        /// </summary>
        public static string GetLogFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_ppt_log.txt");
        }

        /// <summary>
        /// 採点用証跡ログのパス（%TEMP%\mos_ppt_task_evidence.txt）。
        /// 1-2/1-3/1-4/1-8/4-3/5-1/10-4/11-7 など、単体プロジェクトリセット後も採点に必要な行だけを VSTO が追記する。
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

        /// <summary>current_task ファイルの固定フィールド数（ProjectId,TaskId,ExemptFlags,AttemptNo,SnapshotGen）。</summary>
        public const int CurrentTaskFieldCount = 5;

        /// <summary>
        /// 現在タスク共有ファイルを原子的に書き込む。
        /// 形式: ProjectId,TaskId,ExemptFlags,AttemptNo,SnapshotGen（5項目固定）。
        /// UI 遷移は SnapshotGen=0、採点時は &gt;0。
        /// </summary>
        public static void WriteCurrentTaskFile(int projectId, int taskId, int exemptFlags, int attemptNo, int snapshotGen = 0)
        {
            try
            {
                if (attemptNo < 1) attemptNo = 1;
                if (snapshotGen < 0) snapshotGen = 0;
                string content = $"{projectId},{taskId},{exemptFlags},{attemptNo},{snapshotGen}";
                AtomicWriteAllText(GetCurrentTaskFilePath(), content);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] WriteCurrentTaskFile: " + ex.Message);
            }
        }

        /// <summary>
        /// current_task を読み取る。5項目固定かつ全フィールドが数値として解釈できる場合のみ true。
        /// 途中書き込み（フィールド数不足）や空行は無視する。
        /// </summary>
        public static bool TryReadCurrentTaskFile(out int projectId, out int taskId, out int exemptFlags, out int attemptNo, out int snapshotGen)
        {
            return TryReadCurrentTaskFile(GetCurrentTaskFilePath(), out projectId, out taskId, out exemptFlags, out attemptNo, out snapshotGen);
        }

        /// <summary>指定パスの current_task を読み取る（<see cref="TryReadCurrentTaskFile(out int, out int, out int, out int, out int)"/> と同条件）。</summary>
        public static bool TryReadCurrentTaskFile(string path, out int projectId, out int taskId, out int exemptFlags, out int attemptNo, out int snapshotGen)
        {
            projectId = taskId = exemptFlags = attemptNo = snapshotGen = 0;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return false;

            string line;
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var sr = new StreamReader(fs, Encoding.UTF8))
                    line = sr.ReadToEnd().Trim();
            }
            catch
            {
                return false;
            }

            if (string.IsNullOrEmpty(line))
                return false;

            string[] parts = line.Split(',');
            if (parts.Length != CurrentTaskFieldCount)
                return false;

            if (!int.TryParse(parts[0].Trim(), out projectId)) return false;
            if (!int.TryParse(parts[1].Trim(), out taskId)) return false;
            if (!int.TryParse(parts[2].Trim(), out exemptFlags)) return false;
            if (!int.TryParse(parts[3].Trim(), out attemptNo)) return false;
            if (!int.TryParse(parts[4].Trim(), out snapshotGen)) return false;
            if (attemptNo < 1) attemptNo = 1;
            if (snapshotGen < 0) snapshotGen = 0;
            return true;
        }

        private static void AtomicWriteAllText(string path, string content)
        {
            string tempPath = path + ".tmp";
            var encoding = new UTF8Encoding(false);
            File.WriteAllText(tempPath, content, encoding);
            try
            {
                if (File.Exists(path))
                    File.Replace(tempPath, path, null);
                else
                    File.Move(tempPath, path);
            }
            catch
            {
                try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
                throw;
            }
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
        /// 1→1-2/1-3/1-4/1-8、4→4-3、5→5-1、10→10-4、11→11-7。他プロジェクトでは何もしない。
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
                case 1:
                    return new[]
                    {
                        "[Task1-2] Duplicate",
                        "[Task1-3] HideSlide3",
                        "[Task1-4] DeleteThirdSlide",
                        "[Task1-8] SummaryZoom"
                    };
                case 5:
                    return new[] { "[Task5-1] Print" };
                case 4:
                    return new[] { "[Task4-3] Glow18Accent6" };
                case 8:
                    return new[] { "[Task8-3] SlideSize16x9", "[Task8-5] Grayscale" };
                case 9:
                    return new[] { "[Task9-3] PlayAcrossSlides", "[Task9-3] FadeOut3000" };
                case 10:
                    return new[] { "[Task10-4] Grayscale" };
                case 6:
                    return new[] { "[Task6-5] Print", "[Task6-6] Print", "[Task6-7] Print" };
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

        /// <summary>破壊的操作ログに同一 project-task-attempt の記録があるか。</summary>
        public static bool HasLoggedDestructiveError(int projectId, int taskId, int attemptNo)
        {
            try
            {
                string path = GetDestructiveLogPath();
                if (!File.Exists(path)) return false;
                string prefix = $"{projectId},{taskId},{attemptNo}:";
                foreach (string line in File.ReadAllLines(path))
                {
                    if (line != null && line.StartsWith(prefix, StringComparison.Ordinal))
                        return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] HasLoggedDestructiveError: " + ex.Message);
            }
            return false;
        }

        /// <summary>破壊的操作ログへ追記（Word の AppendDestructiveErrors 相当）。</summary>
        public static void AppendDestructiveErrors(int projectId, int taskId, int attemptNo, IList<string> errors)
        {
            if (errors == null || errors.Count == 0) return;
            try
            {
                string body = string.Join(" | ", errors.Where(e => !string.IsNullOrWhiteSpace(e)));
                if (string.IsNullOrWhiteSpace(body)) return;
                string key = $"{projectId},{taskId},{attemptNo}:";
                File.AppendAllText(GetDestructiveLogPath(), key + body + Environment.NewLine, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] AppendDestructiveErrors: " + ex.Message);
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

        private static int _snapshotGenerationCounter;

        /// <summary>採点直前のスナップショット再取得用に、単調増加の世代番号を払い出す。</summary>
        public static int AllocateSnapshotGeneration()
        {
            return Interlocked.Increment(ref _snapshotGenerationCounter);
        }

        /// <summary>スナップショットファイルから TaskId / AttemptNo / SnapshotGen を読み取る。</summary>
        public static bool TryReadSnapshotMeta(string snapshotPath, out int projectId, out int taskId, out int attemptNo, out int snapshotGen)
        {
            projectId = -1;
            taskId = -1;
            attemptNo = 1;
            snapshotGen = 0;
            if (string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
                return false;

            bool hasTaskId = false;
            foreach (string line in File.ReadAllLines(snapshotPath))
            {
                if (string.IsNullOrEmpty(line)) continue;
                int colonIndex = line.IndexOf(':');
                if (colonIndex < 0) continue;
                string key = line.Substring(0, colonIndex);
                string value = line.Substring(colonIndex + 1);
                if (string.Equals(key, "TaskId", StringComparison.Ordinal))
                {
                    var ids = value.Split(',');
                    if (ids.Length != 2) return false;
                    hasTaskId = int.TryParse(ids[0], out projectId) && int.TryParse(ids[1], out taskId);
                }
                else if (string.Equals(key, "SnapshotGen", StringComparison.Ordinal))
                {
                    int.TryParse(value, out snapshotGen);
                }
                else if (string.Equals(key, "AttemptNo", StringComparison.Ordinal))
                {
                    int.TryParse(value, out attemptNo);
                    if (attemptNo < 1) attemptNo = 1;
                }
            }
            return hasTaskId;
        }

        /// <summary>後方互換: AttemptNo を返さないオーバーロード。</summary>
        public static bool TryReadSnapshotMeta(string snapshotPath, out int projectId, out int taskId, out int snapshotGen)
        {
            bool ok = TryReadSnapshotMeta(snapshotPath, out projectId, out taskId, out int attemptNo, out snapshotGen);
            return ok;
        }

        /// <summary>
        /// 証跡ログまたはメインログに [Task10-4] Grayscale が含まれるか（移行前のセッションはメインログのみの可能性あり）。
        /// </summary>
        public static bool HasTask10_4GrayscaleExecuted()
        {
            return HasGradingEvidenceMarker("[Task10-4] Grayscale");
        }

        /// <summary>証跡ログまたはメインログに [Task8-5] Grayscale が含まれるか。</summary>
        public static bool HasTask8_5GrayscaleExecuted()
        {
            return HasGradingEvidenceMarker("[Task8-5] Grayscale");
        }

        /// <summary>証跡ログまたはメインログに [Task8-3] SlideSize16x9 が含まれるか（P8-4で寸法上書き後の一括採点用）。</summary>
        public static bool HasTask8_3SlideSize16x9Executed()
        {
            return HasGradingEvidenceMarker("[Task8-3] SlideSize16x9");
        }

        /// <summary>証跡またはメインログに 4-3 光彩設定記録（[Task4-3] Glow18Accent6）が含まれるか。</summary>
        public static bool HasTask4_3GlowExecuted()
        {
            return HasGradingEvidenceMarker("[Task4-3] Glow18Accent6");
        }

        /// <summary>証跡またはメインログに 5-1 の印刷記録（[Task5-1] Print）が含まれるか。</summary>
        public static bool HasTask5_1PrintExecuted()
        {
            return HasGradingEvidenceMarker("[Task5-1] Print");
        }

        /// <summary>証跡またはメインログに P6-5 の印刷記録（[Task6-5] Print）が含まれるか。</summary>
        public static bool HasTask6_5PrintExecuted()
        {
            return HasGradingEvidenceMarker("[Task6-5] Print");
        }

        /// <summary>証跡またはメインログに P6-6 の印刷記録（[Task6-6] Print）が含まれるか。</summary>
        public static bool HasTask6_6PrintExecuted()
        {
            return HasGradingEvidenceMarker("[Task6-6] Print");
        }

        /// <summary>証跡またはメインログに P6-7 の印刷記録（[Task6-7] Print）が含まれるか。</summary>
        public static bool HasTask6_7PrintExecuted()
        {
            return HasGradingEvidenceMarker("[Task6-7] Print");
        }

        /// <summary>証跡またはメインログに 1-8 サマリーズーム挿入記録（[Task1-8] SummaryZoom）が含まれるか。</summary>
        public static bool HasTask1_8SummaryZoomExecuted()
        {
            return HasGradingEvidenceMarker("[Task1-8] SummaryZoom");
        }

        /// <summary>
        /// セッション内で 1-8 サマリーズームが実行済みか（採点コンテキストに依存しない）。
        /// タスク1-6のスライド番号補正に使用する。
        /// </summary>
        public static bool HasTask1_8SummaryZoomExecutedGlobally()
        {
            return FileContainsMarker(GetTaskEvidenceLogPath(), "[Task1-8] SummaryZoom")
                || FileContainsMarker(GetLogFilePath(), "[Task1-8] SummaryZoom");
        }

        /// <summary>採点用証跡を優先し、無ければ従来の mos_ppt_log.txt を検索する。</summary>
        private static bool HasGradingEvidenceMarker(string marker)
        {
            if (_gradingProjectId.Value.HasValue && _gradingTaskId.Value.HasValue && _gradingAttemptNo.Value.HasValue)
            {
                int p = _gradingProjectId.Value.Value;
                int t = _gradingTaskId.Value.Value;
                int a = _gradingAttemptNo.Value.Value;
                return HasMarkerWithinTask(GetTaskEvidenceLogPath(), p, t, a, marker)
                    || HasMarkerWithinTask(GetLogFilePath(), p, t, a, marker);
            }
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
            if (_gradingProjectId.Value == projectId && _gradingTaskId.Value == taskId && _gradingAttemptNo.Value.HasValue)
                return HasMarkerWithinTask(projectId, taskId, _gradingAttemptNo.Value.Value, marker);
            return HasMarkerWithinTask(projectId, taskId, 1, marker);
        }

        public static bool HasMarkerWithinTask(int projectId, int taskId, int attemptNo, string marker)
        {
            if (string.IsNullOrEmpty(marker))
                return false;
            return HasMarkerWithinTask(GetLogFilePath(), projectId, taskId, attemptNo, marker);
        }

        private static bool HasMarkerWithinTask(string path, int projectId, int taskId, int attemptNo, string marker)
        {
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(marker) || !File.Exists(path))
                return false;
            try
            {
                string[] lines = File.ReadAllLines(path);
                bool inTarget = false;
                int curP = -1, curT = -1, curA = 1;
                foreach (string line in lines)
                {
                    if (line == null) continue;
                    string taskPrefix = $"[Task {projectId}-{taskId}-{attemptNo}]";
                    if (line.IndexOf(taskPrefix, StringComparison.OrdinalIgnoreCase) >= 0
                        && line.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;

                    if (line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ParseTaskStart(line, out curP, out curT, out curA);
                        inTarget = (curP == projectId && curT == taskId && curA == attemptNo);
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

        /// <summary>ログに P6-3 Kiosk 設定記録（[Task6-3] Kiosk）が含まれるか。</summary>
        public static bool HasTask6_3KioskExecuted()
        {
            return HasLogLineContaining("[Task6-3] Kiosk");
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
            return GetOperationsForTask(projectId, taskId, 1);
        }

        public static List<string> GetOperationsForTask(int projectId, int taskId, int attemptNo)
        {
            var result = new List<string>();
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return result;
            try
            {
                string[] lines = File.ReadAllLines(path);
                int currentProject = -1, currentTask = -1, currentAttempt = 1;
                foreach (string line in lines)
                {
                    if (line == null) continue;
                    if (line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ParseTaskStart(line, out currentProject, out currentTask, out currentAttempt);
                        continue;
                    }
                    if (currentProject != projectId || currentTask != taskId || currentAttempt != attemptNo)
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
            return HasDisallowedOperations(projectId, taskId, 1, allowedOperationTypes);
        }

        public static bool HasDisallowedOperations(int projectId, int taskId, int attemptNo, HashSet<string> allowedOperationTypes)
        {
            if (allowedOperationTypes == null)
                return false;
            var ops = GetOperationsForTask(projectId, taskId, attemptNo);
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
            var ops = GetOperationsForTask(projectId, taskId, 1);
            foreach (string opLine in ops)
            {
                if (opLine.IndexOf("ShapePositionChange", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        private static void ParseTaskStart(string line, out int projectId, out int taskId, out int attemptNo)
        {
            projectId = -1;
            taskId = -1;
            attemptNo = 1;
            int startIdx = line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase);
            if (startIdx < 0) return;
            string part = line.Substring(startIdx + 11).Trim();
            var tokens = part.Split(new[] { '-', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length >= 2 && int.TryParse(tokens[0].Trim(), out projectId) && int.TryParse(tokens[1].Trim(), out taskId))
            {
                if (tokens.Length >= 3)
                {
                    int.TryParse(tokens[2].Trim(), out attemptNo);
                    if (attemptNo < 1) attemptNo = 1;
                }
                return;
            }
            projectId = -1;
            taskId = -1;
            attemptNo = 1;
        }


        private static string GetOperationType(string opLine)
        {
            if (string.IsNullOrWhiteSpace(opLine)) return "";
            int space = opLine.IndexOf(' ');
            return space > 0 ? opLine.Substring(0, space).Trim() : opLine.Trim();
        }

        /// <summary>タスク開始時スナップショットの読み取り結果（スライド構成検証用）。</summary>
        public sealed class PPTaskSnapshotData
        {
            public int ProjectId { get; set; }
            public int TaskId { get; set; }
            public int SlidesCount { get; set; }
            public List<string> SlideNames { get; set; } = new List<string>();
        }

        /// <summary>
        /// %TEMP%\mos_ppt_snapshot.txt から、指定タスクの開始時スナップショットを読み込む。
        /// ProjectId / TaskId が一致しない場合は false を返す。
        /// </summary>
        public static bool TryLoadTaskSnapshot(int projectId, int taskId, out PPTaskSnapshotData snapshot)
        {
            snapshot = null;
            string path = GetSnapshotPath();
            if (!File.Exists(path))
                return false;

            try
            {
                var data = new PPTaskSnapshotData();
                foreach (string line in File.ReadAllLines(path))
                {
                    if (string.IsNullOrEmpty(line)) continue;
                    int colonIndex = line.IndexOf(':');
                    if (colonIndex < 0) continue;

                    string key = line.Substring(0, colonIndex);
                    string value = line.Substring(colonIndex + 1);

                    switch (key)
                    {
                        case "TaskId":
                            var ids = value.Split(',');
                            if (ids.Length == 2)
                            {
                                int.TryParse(ids[0], out int pid);
                                int.TryParse(ids[1], out int tid);
                                data.ProjectId = pid;
                                data.TaskId = tid;
                            }
                            break;
                        case "SlidesCount":
                            int.TryParse(value, out int slidesCount);
                            data.SlidesCount = slidesCount;
                            break;
                        case "SlideNames":
                            data.SlideNames = value.Split(new[] { '|' }, StringSplitOptions.None).ToList();
                            break;
                    }
                }

                if (data.ProjectId != projectId || data.TaskId != taskId)
                    return false;
                if (data.SlidesCount < 1 || data.SlideNames == null || data.SlideNames.Count != data.SlidesCount)
                    return false;

                snapshot = data;
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
