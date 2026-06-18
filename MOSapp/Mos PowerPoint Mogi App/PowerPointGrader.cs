using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Interop.PowerPoint;
using Libraries;
using Libraries.Group1;

namespace MOS_PowerPoint_app
{
    /// <summary>
    /// 起動中の PowerPoint に接続し、MOS 模擬試験のタスクごとの採点を行うクラス。
    /// 採点ロジックは Libraries.Group1 の PowerPointChecker1_X に委譲する。
    /// </summary>
    public sealed class PowerPointGrader : IDisposable
    {
        private Application _app;
        private Presentation _activePresentation;
        private bool _disposed;

        /// <summary>
        /// 現在起動している PowerPoint インスタンスに接続する。
        /// </summary>
        /// <returns>接続に成功した場合は true、PowerPoint が起動していない等で失敗した場合は false。</returns>
        public bool Connect()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                try
                {
                    if (_app == null)
                    {
                        _app = (Application)Marshal.GetActiveObject("PowerPoint.Application");
                        if (_app == null)
                            return false;
                    }

                    // プレゼンを閉じて別ファイルを開いた直後など、古い Presentation 参照は無効になる。
                    // 毎回 ActivePresentation を取り直す（全プロジェクト連続採点で必須）。
                    if (_activePresentation != null)
                    {
                        try { Marshal.ReleaseComObject(_activePresentation); } catch { }
                        _activePresentation = null;
                    }

                    try
                    {
                        _activePresentation = _app.ActivePresentation;
                    }
                    catch
                    {
                        _activePresentation = null;
                    }

                    return _activePresentation != null;
                }
                catch (COMException)
                {
                    _app = null;
                    _activePresentation = null;
                    return false;
                }
            }
        }

        /// <summary>
        /// 指定したプロジェクト・タスクの採点を行う。
        /// ログに余計な操作や許可されない座標変化があれば不合格。続けて COM による結果判定を行う。
        /// </summary>
        /// <param name="projectId">プロジェクト ID（1～11）。</param>
        /// <param name="taskId">タスク ID。</param>
        /// <returns>合格なら true、不合格または未実装・範囲外なら false。</returns>
        public void StartTask(int projectId, int taskId, int attemptNo = 1)
        {
            try
            {
                var flags = Libraries.PPTaskValidationConfig.GetExemptFlags(projectId, taskId);
                if (attemptNo < 1) attemptNo = 1;
                File.WriteAllText(Libraries.PPLogReader.GetCurrentTaskFilePath(), $"{projectId},{taskId},{(int)flags},{attemptNo}");
            }
            catch { }
        }

        /// <summary>
        /// StartTask の書き込み後、VSTO 側のスナップショット（%TEMP%\mos_ppt_snapshot.txt）が
        /// 指定の projectId-taskId に更新されるまで短時間待機する。
        /// タイムアウトした場合は待機を諦め、採点は継続する（スナップショットチェックは ID Mismatch でスキップされる）。
        /// </summary>
        public void StartTaskAndWaitForSnapshot(int projectId, int taskId, int timeoutMs = 2000, int pollIntervalMs = 50)
        {
            StartTaskAndWaitForSnapshot(projectId, taskId, 1, timeoutMs, pollIntervalMs);
        }

        public void StartTaskAndWaitForSnapshot(int projectId, int taskId, int attemptNo, int timeoutMs = 2000, int pollIntervalMs = 50)
        {
            var swTotal = Stopwatch.StartNew();
            StartTask(projectId, taskId, attemptNo);

            // VSTO 側は一定間隔（例: 500ms）で current_task を監視して snapshot を更新するため、短時間だけ待つ。
            string snapshotPath = Libraries.PPLogReader.GetSnapshotPath();
            var swWait = Stopwatch.StartNew();
            while (swWait.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    if (TryReadSnapshotTaskId(snapshotPath, out int snapProjectId, out int snapTaskId))
                    {
                        if (snapProjectId == projectId && snapTaskId == taskId)
                        {
                            PPGradingPerf.Log("StartTaskAndWaitForSnapshot.wait", swWait.ElapsedMilliseconds, $"P{projectId}-T{taskId} snapshot matched");
                            PPGradingPerf.Log("StartTaskAndWaitForSnapshot.total", swTotal.ElapsedMilliseconds, $"P{projectId}-T{taskId}");
                            return;
                        }
                    }
                }
                catch { }
                Thread.Sleep(pollIntervalMs);
            }
            PPGradingPerf.Log("StartTaskAndWaitForSnapshot.wait", swWait.ElapsedMilliseconds, $"P{projectId}-T{taskId} timeout {timeoutMs}ms");
            PPGradingPerf.Log("StartTaskAndWaitForSnapshot.total", swTotal.ElapsedMilliseconds, $"P{projectId}-T{taskId}");
        }

        private static bool TryReadSnapshotTaskId(string snapshotPath, out int projectId, out int taskId)
        {
            projectId = -1;
            taskId = -1;
            if (string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
                return false;

            // 期待フォーマット例: "TaskId:1,4"
            // PPSnapshotChecker.LoadSnapshot と同じ形式を読む。
            string[] lines = File.ReadAllLines(snapshotPath);
            foreach (string line in lines)
            {
                if (string.IsNullOrEmpty(line)) continue;
                int colonIndex = line.IndexOf(':');
                if (colonIndex < 0) continue;
                string key = line.Substring(0, colonIndex);
                if (!string.Equals(key, "TaskId", StringComparison.Ordinal)) continue;
                string value = line.Substring(colonIndex + 1);
                var ids = value.Split(',');
                if (ids.Length != 2) return false;
                return int.TryParse(ids[0], out projectId) && int.TryParse(ids[1], out taskId);
            }
            return false;
        }
        /// <summary>
        /// 指定したプロジェクト・タスクの採点を行う。
        /// ログに余計な操作や許可されない座標変化があれば不合格。続けて COM による結果判定を行う。
        /// </summary>
        /// <param name="projectId">プロジェクト ID（1～10）。</param>
        /// <param name="taskId">タスク ID。</param>
        /// <returns>合格なら true、不合格または未実装・範囲外なら false。</returns>
        public bool GradeTask(int projectId, int taskId)
        {
            return GradeTask(projectId, taskId, 1);
        }

        public bool GradeTask(int projectId, int taskId, int attemptNo)
        {
            var swGradeTotal = Stopwatch.StartNew();
            if (_activePresentation == null)
                return false;

            var sw = Stopwatch.StartNew();
            // 1. 過去の破壊的操作ログのチェック
            if (HasLoggedDestructiveError(projectId, taskId, attemptNo))
            {
                System.Diagnostics.Debug.WriteLine($"[Grader] Task {projectId}-{taskId} FAILED due to logged destructive operation.");
                if (projectId == 1 && taskId == 1)
                    Debug.WriteLine("[Task1-1] GradeTask: FAIL early exit (logged destructive operation)");
                PPGradingPerf.Log("GradeTask.HasLoggedDestructiveError", sw.ElapsedMilliseconds, $"P{projectId}-T{taskId}");
                PPGradingPerf.Log("GradeTask.total", swGradeTotal.ElapsedMilliseconds, $"P{projectId}-T{taskId} early exit");
                return false;
            }
            PPGradingPerf.Log("GradeTask.HasLoggedDestructiveError", sw.ElapsedMilliseconds, $"P{projectId}-T{taskId}");

            sw.Restart();
            if (FailsLogChecks(projectId, taskId, attemptNo))
            {
                if (projectId == 1 && taskId == 1)
                    Debug.WriteLine("[Task1-1] GradeTask: FAIL early exit (disallowed log operations)");
                PPGradingPerf.Log("GradeTask.FailsLogChecks", sw.ElapsedMilliseconds, $"P{projectId}-T{taskId} failed");
                PPGradingPerf.Log("GradeTask.total", swGradeTotal.ElapsedMilliseconds, $"P{projectId}-T{taskId} early exit");
                return false;
            }
            PPGradingPerf.Log("GradeTask.FailsLogChecks", sw.ElapsedMilliseconds, $"P{projectId}-T{taskId}");

            // 2. スナップショット比較と COM 採点を同一ロックで実行（閉じる処理とスナップショット COM の間に割り込まれないようにする）
            sw.Restart();
            var exemptFlags = Libraries.PPTaskValidationConfig.GetExemptFlags(projectId, taskId);
            bool comResult = false;
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                var destructiveErrors = Libraries.PPSnapshotChecker.CompareAndGetErrors(projectId, taskId, exemptFlags);
                PPGradingPerf.Log("GradeTask.PPSnapshotCompare", sw.ElapsedMilliseconds, $"P{projectId}-T{taskId}");
                if (projectId == 1 && taskId == 1)
                {
                    LogTask1_1SnapshotContext(destructiveErrors);
                }
                if (destructiveErrors.Count > 0)
                {
                    foreach (var err in destructiveErrors)
                    {
                        System.Diagnostics.Debug.WriteLine($"[Validation] Project{projectId} Task{taskId}: {err}");
                    }
                    if (projectId == 1 && taskId == 1)
                    {
                        Debug.WriteLine($"[Task1-1] GradeTask: FAIL early exit (snapshot errors={destructiveErrors.Count}, COM not run)");
                    }
                    PPGradingPerf.Log("GradeTask.total", swGradeTotal.ElapsedMilliseconds, $"P{projectId}-T{taskId} early exit snapshot errors");
                    return false;
                }

                if (projectId == 1 && taskId == 1)
                {
                    Debug.WriteLine("[Task1-1] GradeTask: snapshot OK, reaching COM checker");
                }

                sw.Restart();
                try
                {
                    PPLogReader.SetGradingContext(projectId, taskId, attemptNo);
                    comResult = RunComChecker(projectId, taskId);
                }
                catch
                {
                    comResult = false;
                }
                finally
                {
                    PPLogReader.ClearGradingContext();
                }
                PPGradingPerf.Log("GradeTask.ComChecker", sw.ElapsedMilliseconds, $"P{projectId}-T{taskId} pass={comResult}");
                if (projectId == 1 && taskId == 1)
                    Debug.WriteLine($"[Task1-1] GradeTask: COM result={(comResult ? "PASS" : "FAIL")}");
            }
            PPGradingPerf.Log("GradeTask.total", swGradeTotal.ElapsedMilliseconds, $"P{projectId}-T{taskId} pass={comResult}");
            return comResult;
        }

        private static void LogTask1_1SnapshotContext(List<string> destructiveErrors)
        {
            PPLogReader.PPTaskSnapshotData snap;
            bool hasSnap = PPLogReader.TryLoadTaskSnapshot(1, 1, out snap);
            string snapInfo = hasSnap
                ? $"SlidesCount={snap.SlidesCount} SlideNames={snap.SlideNames?.Count ?? 0}"
                : "not loaded or ID mismatch";
            Debug.WriteLine($"[Task1-1] GradeTask: snapshot context ({snapInfo}, errors={destructiveErrors.Count})");
            foreach (string err in destructiveErrors)
                Debug.WriteLine($"[Task1-1] GradeTask: snapshot error: {err}");
        }

        /// <summary>プロジェクト別の COM 採点のみ（計測用に分離）。</summary>
        private static bool RunComChecker(int projectId, int taskId)
        {
            switch (projectId)
            {
                case 1:
                    var c1 = new PowerPointChecker1_1();
                    switch (taskId)
                    {
                        case 1: return c1.CheckTask_1_1_01();
                        case 2: return c1.CheckTask_1_1_02();
                        case 3: return c1.CheckTask_1_1_03();
                        case 4: return c1.CheckTask_1_1_04();
                        case 5: return c1.CheckTask_1_1_05();
                        case 6: return c1.CheckTask_1_1_06();
                        case 7: return c1.CheckTask_1_1_07();
                        case 8: return c1.CheckTask_1_1_08();
                        default: return false;
                    }
                case 2:
                    var c2 = new PowerPointChecker1_2();
                    switch (taskId)
                    {
                        case 1: return c2.CheckTask_1_2_01();
                        case 2: return c2.CheckTask_1_2_02();
                        case 3: return c2.CheckTask_1_2_03();
                        case 4: return c2.CheckTask_1_2_04();
                        case 5: return c2.CheckTask_1_2_05();
                        case 6: return c2.CheckTask_1_2_06();
                        case 7: return c2.CheckTask_1_2_07();
                        case 8: return c2.CheckTask_1_2_08();
                        default: return false;
                    }
                case 3:
                    var c3 = new PowerPointChecker1_3();
                    switch (taskId)
                    {
                        case 1: return c3.CheckTask_1_3_01();
                        case 2: return c3.CheckTask_1_3_02();
                        case 3: return c3.CheckTask_1_3_03();
                        case 4: return c3.CheckTask_1_3_04();
                        case 5: return c3.CheckTask_1_3_05();
                        case 6: return c3.CheckTask_1_3_06();
                        case 7: return c3.CheckTask_1_3_07();
                        default: return false;
                    }
                case 4:
                    var c4 = new PowerPointChecker1_4();
                    switch (taskId)
                    {
                        case 1: return c4.CheckTask_1_4_01();
                        case 2: return c4.CheckTask_1_4_02();
                        case 3: return c4.CheckTask_1_4_03();
                        case 4: return c4.CheckTask_1_4_04();
                        case 5: return c4.CheckTask_1_4_05();
                        case 6: return c4.CheckTask_1_4_06();
                        case 7: return c4.CheckTask_1_4_07();
                        case 8: return c4.CheckTask_1_4_08();
                        default: return false;
                    }
                case 5:
                    var c5 = new PowerPointChecker1_5();
                    switch (taskId)
                    {
                        case 1: return c5.CheckTask_1_5_01();
                        case 2: return c5.CheckTask_1_5_02();
                        case 3: return c5.CheckTask_1_5_03();
                        case 4: return c5.CheckTask_1_5_04();
                        case 5: return c5.CheckTask_1_5_05();
                        default: return false;
                    }
                case 6:
                    var c6 = new PowerPointChecker1_6();
                    switch (taskId)
                    {
                        case 1: return c6.CheckTask_1_6_01();
                        case 2: return c6.CheckTask_1_6_02();
                        case 3: return c6.CheckTask_1_6_03();
                        case 4: return c6.CheckTask_1_6_04();
                        default: return false;
                    }
                case 7:
                    var c7 = new PowerPointChecker1_7();
                    switch (taskId)
                    {
                        case 1: return c7.CheckTask_1_7_01();
                        case 2: return c7.CheckTask_1_7_02();
                        case 3: return c7.CheckTask_1_7_03();
                        case 4: return c7.CheckTask_1_7_04();
                        default: return false;
                    }
                case 8:
                    var c8 = new PowerPointChecker1_8();
                    switch (taskId)
                    {
                        case 1: return c8.CheckTask_1_8_01();
                        case 2: return c8.CheckTask_1_8_02();
                        case 3: return c8.CheckTask_1_8_03();
                        case 4: return c8.CheckTask_1_8_04();
                        case 5: return c8.CheckTask_1_8_05();
                        default: return false;
                    }
                case 9:
                    var c9 = new PowerPointChecker1_9();
                    switch (taskId)
                    {
                        case 1: return c9.CheckTask_1_9_01();
                        case 2: return c9.CheckTask_1_9_02();
                        case 3: return c9.CheckTask_1_9_03();
                        case 4: return c9.CheckTask_1_9_04();
                        case 5: return c9.CheckTask_1_9_05();
                        case 6: return c9.CheckTask_1_9_06();
                        case 7: return c9.CheckTask_1_9_07();
                        default: return false;
                    }
                case 10:
                    var c10 = new PowerPointChecker1_10();
                    switch (taskId)
                    {
                        case 1: return c10.CheckTask_1_10_01();
                        case 2: return c10.CheckTask_1_10_02();
                        case 3: return c10.CheckTask_1_10_03();
                        case 4: return c10.CheckTask_1_10_04();
                        case 5: return c10.CheckTask_1_10_05();
                        case 6: return c10.CheckTask_1_10_06();
                        case 7: return c10.CheckTask_1_10_07();
                        default: return false;
                    }
                default:
                    return false;
            }
        }

        /// <summary>
        /// VSTO ログを参照し、余計な操作または許可されない座標変化があれば true（不合格とする）。
        /// ログファイルが無い場合は false（アドイン未導入時は COM のみで判定）。
        /// </summary>
        private static readonly string DestructiveLogPath = Path.Combine(Path.GetTempPath(), "mos_ppt_destructive_errors.log");

        private bool HasLoggedDestructiveError(int projectId, int taskId, int attemptNo)
        {
            try
            {
                if (!File.Exists(DestructiveLogPath)) return false;
                var lines = File.ReadAllLines(DestructiveLogPath);
                string prefix = $"{projectId},{taskId},{attemptNo}:";
                string legacyPrefix = $"{projectId},{taskId}:";
                foreach (var line in lines)
                {
                    if (line.StartsWith(prefix)) return true;
                    if (attemptNo <= 1 && line.StartsWith(legacyPrefix)) return true;
                }
            }
            catch { }
            return false;
        }

        private static bool FailsLogChecks(int projectId, int taskId, int attemptNo)
        {
            string logPath = PPLogReader.GetLogFilePath();
            if (!File.Exists(logPath))
                return false;

            // [Op] は旧リボン上書きで RibbonCommand のみ出力。上書き廃止後は通常該当なし。将来 LogOperation を増やす場合は allowed を調整。
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "RibbonCommand" };
            if (PPLogReader.HasDisallowedOperations(projectId, taskId, attemptNo, allowed))
                return true;
            
            return false;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (_disposed)
                return;
            if (disposing)
            {
                lock (PowerPointCheckerCommon.PowerPointComInteropSync)
                {
                    try
                    {
                        if (_activePresentation != null)
                        {
                            Marshal.ReleaseComObject(_activePresentation);
                            _activePresentation = null;
                        }
                    }
                    catch { }
                    try
                    {
                        if (_app != null)
                        {
                            Marshal.ReleaseComObject(_app);
                            _app = null;
                        }
                    }
                    catch { }
                }
            }
            _disposed = true;
        }
    }
}
