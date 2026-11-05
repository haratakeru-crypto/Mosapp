using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Libraries;
using Libraries.Group1;

namespace MOS_PowerPoint_app
{
    /// <summary>
    /// 類題セット1のその場採点。実装済み PV チェッカーを呼び、未実装プロジェクトは不合格扱い。
    /// Word の WordGradingGate / 教材 GradeTask と同様、COM 前に破壊的ゲートを通す。
    /// </summary>
    public static class PowerPointVariantScoring
    {
        public const string SetSuffix = "PV1";
        public const int DefaultVariantSetNo = 1;

        private static readonly HashSet<string> AllowedOperationTypes =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "RibbonCommand" };

        public static IReadOnlyList<TaskResult> ScoreProject(int groupId, int projectId, int variantSetNo = DefaultVariantSetNo)
        {
            if (variantSetNo < 1)
                variantSetNo = DefaultVariantSetNo;

            var tasks = VariantProblemWorkbook.GetTasks(projectId);
            var results = new List<TaskResult>(tasks.Count);
            if (tasks.Count == 0)
                return results;

            // 類題 VSTO 証跡（TaskV1-*）を採点直前に確定させる。
            try { PPLogReader.WriteVariantSetNo(variantSetNo); } catch { }
            try { PPLogReader.RequestOpenTaskEvidenceFlush(projectId); } catch { }

            object checker = TryCreateChecker(groupId, projectId);
            foreach (VariantTaskText task in tasks)
            {
                bool passed = false;
                try
                {
                    int attemptNo = PPTaskAttemptRegistry.GetAttempt(projectId, task.TaskId, variantSetNo);
                    // 採点は Task.Run から呼ばれるため、PowerPoint COM を直列化する。
                    lock (PowerPointCheckerCommon.PowerPointComInteropSync)
                    {
                        if (!TryPassDestructiveGate(projectId, task.TaskId, attemptNo, variantSetNo))
                        {
                            Debug.WriteLine(
                                $"[PowerPointVariantScoring] P{projectId}-T{task.TaskId}-A{attemptNo} V{variantSetNo}: FAIL destructive gate");
                            passed = false;
                        }
                        else
                        {
                            passed = TryScoreTask(checker, groupId, projectId, task.TaskId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerPointVariantScoring] P{projectId}-T{task.TaskId}: {ex.Message}");
                    passed = false;
                }

                results.Add(new TaskResult
                {
                    TaskNumber = task.TaskId,
                    IsPassed = passed,
                    TaskName = string.IsNullOrEmpty(task.Description)
                        ? $"タスク{task.TaskId}"
                        : task.Description
                });

                ScoreResultStore.RecordResult(groupId, projectId, task.TaskId, passed, null, variantSetNo);
            }

            return results;
        }

        /// <summary>
        /// 教材 GradeTask / WordGradingGate 相当。
        /// ① destructive ログ ② 許可外 [Op] ③ スナップショット差分（一致タスクのみ）。
        /// </summary>
        private static bool TryPassDestructiveGate(int projectId, int taskId, int attemptNo, int variantSetNo)
        {
            if (PPLogReader.HasLoggedDestructiveError(projectId, taskId, attemptNo, variantSetNo))
                return false;

            if (PPLogReader.HasDisallowedOperations(projectId, taskId, attemptNo, AllowedOperationTypes))
                return false;

            var exemptFlags = PPVariantValidationConfig.ResolveExemptFlags(variantSetNo, projectId, taskId);
            List<string> errors = PPSnapshotChecker.CompareAndGetErrors(projectId, taskId, attemptNo, exemptFlags);
            if (errors != null && errors.Count > 0)
            {
                foreach (string err in errors)
                    Debug.WriteLine($"[PowerPointVariantScoring] V{variantSetNo} P{projectId}-T{taskId}: {err}");
                PPLogReader.AppendDestructiveErrors(projectId, taskId, attemptNo, errors, variantSetNo);
                return false;
            }

            return true;
        }

        public static bool IsTaskImplemented(int groupId, int projectId, int taskId)
        {
            Type checkerType = GetCheckerType(groupId, projectId);
            if (checkerType == null)
                return false;
            return checkerType.GetMethod(GetMethodName(groupId, projectId, taskId),
                       BindingFlags.Instance | BindingFlags.Public) != null;
        }

        private static bool TryScoreTask(object checker, int groupId, int projectId, int taskId)
        {
            if (checker == null)
                return false;

            MethodInfo method = checker.GetType().GetMethod(
                GetMethodName(groupId, projectId, taskId),
                BindingFlags.Instance | BindingFlags.Public);
            if (method == null)
                return false;

            object value = method.Invoke(checker, null);
            return value is bool flag && flag;
        }

        private static object TryCreateChecker(int groupId, int projectId)
        {
            // 実装済み PV はプロジェクト参照で直接生成（COM 相互運用を安定させる）。
            if (groupId == 1 && projectId == 1)
            {
                try { return new PowerPointChecker1_1_PV1(); }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerPointVariantScoring] P1_PV1 create failed: {ex.Message}");
                    return null;
                }
            }
            if (groupId == 1 && projectId == 2)
            {
                try { return new PowerPointChecker1_2_PV1(); }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerPointVariantScoring] P2_PV1 create failed: {ex.Message}");
                    return null;
                }
            }
            if (groupId == 1 && projectId == 3)
            {
                try { return new PowerPointChecker1_3_PV1(); }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerPointVariantScoring] P3_PV1 create failed: {ex.Message}");
                    return null;
                }
            }
            if (groupId == 1 && projectId == 4)
            {
                try { return new PowerPointChecker1_4_PV1(); }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerPointVariantScoring] P4_PV1 create failed: {ex.Message}");
                    return null;
                }
            }
            if (groupId == 1 && projectId == 5)
            {
                try { return new PowerPointChecker1_5_PV1(); }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerPointVariantScoring] P5_PV1 create failed: {ex.Message}");
                    return null;
                }
            }
            if (groupId == 1 && projectId == 6)
            {
                try { return new PowerPointChecker1_6_PV1(); }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerPointVariantScoring] P6_PV1 create failed: {ex.Message}");
                    return null;
                }
            }

            Type checkerType = GetCheckerType(groupId, projectId);
            if (checkerType == null)
                return null;
            try
            {
                return Activator.CreateInstance(checkerType);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[PowerPointVariantScoring] create failed: {ex.Message}");
                return null;
            }
        }

        private static Type GetCheckerType(int groupId, int projectId)
        {
            string typeName = $"Libraries.Group{groupId}.PowerPointChecker{groupId}_{projectId}_{SetSuffix}";
            return AppDomain.CurrentDomain.GetAssemblies()
                .Select(a =>
                {
                    try { return a.GetType(typeName, throwOnError: false); }
                    catch { return null; }
                })
                .FirstOrDefault(t => t != null);
        }

        private static string GetMethodName(int groupId, int projectId, int taskId)
        {
            return $"CheckTask_{groupId}_{projectId}_{taskId:D2}";
        }
    }
}
