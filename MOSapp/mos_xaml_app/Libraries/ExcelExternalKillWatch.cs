using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace Libraries
{
    /// <summary>
    /// MOS 本体とは別プロセスで Excel PID を遅延 Kill する。
    /// 本体が COM で固まっても、即終了しても、監視プロセスが残れば落とせる。
    /// </summary>
    public static class ExcelExternalKillWatch
    {
        const string ExeName = "MosPracticeSessionWatch.exe";
        const string ModeArg = "kill-excel";

        static string ExePath
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ExeName); }
        }

        /// <summary>
        /// 保存・Close 完了後に呼ぶ。delayMs 後に同一 Excel PID が残っていれば外部プロセスが Kill する。
        /// </summary>
        public static bool TryArm(int pid, DateTime? startTime, int delayMs, string logContext = null)
        {
            if (pid <= 0)
                return false;

            string ctx = string.IsNullOrEmpty(logContext) ? "external-watch" : logContext;
            try
            {
                if (!File.Exists(ExePath))
                {
                    ExcelVstoReadiness.RecordHostEvent(
                        "excel-exit external-watch missing exe path=" + ExePath + " context=" + ctx);
                    return false;
                }

                long ticks = 0;
                if (startTime.HasValue)
                {
                    try { ticks = startTime.Value.ToUniversalTime().Ticks; }
                    catch { ticks = 0; }
                }

                if (delayMs < 0)
                    delayMs = 0;
                if (delayMs > 60000)
                    delayMs = 60000;

                var psi = new ProcessStartInfo
                {
                    FileName = ExePath,
                    Arguments = string.Format(
                        CultureInfo.InvariantCulture,
                        "{0} {1} {2} {3}",
                        ModeArg,
                        pid,
                        ticks,
                        delayMs),
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                Process.Start(psi);
                ExcelVstoReadiness.RecordHostEvent(
                    "excel-exit external-watch armed pid=" + pid
                    + " delayMs=" + delayMs
                    + " startTicks=" + ticks
                    + " context=" + ctx);
                return true;
            }
            catch (Exception ex)
            {
                ExcelVstoReadiness.RecordHostEvent(
                    "excel-exit external-watch failed pid=" + pid
                    + " context=" + ctx
                    + " error=" + ex.GetType().Name + ":" + ex.Message);
                return false;
            }
        }
    }
}
