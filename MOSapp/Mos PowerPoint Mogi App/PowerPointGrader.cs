using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
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
            if (_app != null)
                return true;

            try
            {
                _app = (Application)Marshal.GetActiveObject("PowerPoint.Application");
                if (_app == null)
                    return false;

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

        /// <summary>
        /// 指定したプロジェクト・タスクの採点を行う。
        /// ログに余計な操作や許可されない座標変化があれば不合格。続けて COM による結果判定を行う。
        /// </summary>
        /// <param name="projectId">プロジェクト ID（1～11）。</param>
        /// <param name="taskId">タスク ID。</param>
        /// <returns>合格なら true、不合格または未実装・範囲外なら false。</returns>
        public void StartTask(int projectId, int taskId)
        {
            try
            {
                var flags = Libraries.PPTaskValidationConfig.GetExemptFlags(projectId, taskId);
                File.WriteAllText(Libraries.PPLogReader.GetCurrentTaskFilePath(), $"{projectId},{taskId},{(int)flags}");
            }
            catch { }
        }
        /// <summary>
        /// 指定したプロジェクト・タスクの採点を行う。
        /// ログに余計な操作や許可されない座標変化があれば不合格。続けて COM による結果判定を行う。
        /// </summary>
        /// <param name="projectId">プロジェクト ID（1～11）。</param>
        /// <param name="taskId">タスク ID。</param>
        /// <returns>合格なら true、不合格または未実装・範囲外なら false。</returns>
        public bool GradeTask(int projectId, int taskId)
        {
            if (_activePresentation == null)
                return false;

            // 1. 過去の破壊的操作ログのチェック
            if (HasLoggedDestructiveError(projectId, taskId))
            {
                System.Diagnostics.Debug.WriteLine($"[Grader] Task {projectId}-{taskId} FAILED due to logged destructive operation.");
                return false;
            }

            if (FailsLogChecks(projectId, taskId))
                return false;

            // 2. 現在の破壊的操作（リアルタイムスナップショット）のチェック
            var exemptFlags = Libraries.PPTaskValidationConfig.GetExemptFlags(projectId, taskId);
            var destructiveErrors = Libraries.PPSnapshotChecker.CompareAndGetErrors(projectId, taskId, exemptFlags);
            if (destructiveErrors.Count > 0)
            {
                foreach (var err in destructiveErrors)
                {
                    System.Diagnostics.Debug.WriteLine($"[Validation] Project{projectId} Task{taskId}: {err}");
                }
                return false;
            }

            try
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
                    case 11:
                        var c11 = new PowerPointChecker1_11();
                        switch (taskId)
                        {
                            case 1: return c11.CheckTask_1_11_01();
                            case 2: return c11.CheckTask_1_11_02();
                            case 3: return c11.CheckTask_1_11_03();
                            case 4: return c11.CheckTask_1_11_04();
                            case 5: return c11.CheckTask_1_11_05();
                            case 6: return c11.CheckTask_1_11_06();
                            case 7: return c11.CheckTask_1_11_07();
                            default: return false;
                        }
                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// VSTO ログを参照し、余計な操作または許可されない座標変化があれば true（不合格とする）。
        /// ログファイルが無い場合は false（アドイン未導入時は COM のみで判定）。
        /// </summary>
        private static readonly string DestructiveLogPath = Path.Combine(Path.GetTempPath(), "mos_ppt_destructive_errors.log");

        private bool HasLoggedDestructiveError(int projectId, int taskId)
        {
            try
            {
                if (!File.Exists(DestructiveLogPath)) return false;
                var lines = File.ReadAllLines(DestructiveLogPath);
                string prefix = $"{projectId},{taskId}:";
                foreach (var line in lines)
                {
                    if (line.StartsWith(prefix)) return true;
                }
            }
            catch { }
            return false;
        }

        private static bool FailsLogChecks(int projectId, int taskId)
        {
            string logPath = PPLogReader.GetLogFilePath();
            if (!File.Exists(logPath))
                return false;

            // [Op] は旧リボン上書きで RibbonCommand のみ出力。上書き廃止後は通常該当なし。将来 LogOperation を増やす場合は allowed を調整。
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "RibbonCommand" };
            if (PPLogReader.HasDisallowedOperations(projectId, taskId, allowed))
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
            _disposed = true;
        }
    }
}
