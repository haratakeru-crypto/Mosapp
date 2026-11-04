using System;
using System.IO;

namespace ExcelAddIn1
{
    /// <summary>
    /// Libraries.ExcelDiagFlags と同じ判定。VSTO は本体アセンブリを参照しないため、こちらに置く。
    /// Debug では常に詳細。Release／製品では %TEMP%\mos_excel_diag.on があるときだけ。
    /// </summary>
    internal static class ExcelDiagFlags
    {
        public static bool IsVerbose
        {
            get
            {
#if DEBUG
                return true;
#else
                return FlagFileExists();
#endif
            }
        }

#if !DEBUG
        const int RefreshMs = 5000;
        static readonly object Sync = new object();
        static readonly string FlagPath = Path.Combine(Path.GetTempPath(), "mos_excel_diag.on");
        static bool _cached;
        static int _cachedAt;

        static bool FlagFileExists()
        {
            lock (Sync)
            {
                int now = Environment.TickCount;
                if (_cachedAt != 0 && unchecked(now - _cachedAt) >= 0 && unchecked(now - _cachedAt) < RefreshMs)
                    return _cached;
                try
                {
                    _cached = File.Exists(FlagPath);
                }
                catch
                {
                    _cached = false;
                }
                _cachedAt = now;
                return _cached;
            }
        }
#endif

        public static bool ShouldWriteDiagnostic(string message)
        {
            if (IsVerbose)
                return true;
            if (string.IsNullOrEmpty(message))
                return false;
            return message.IndexOf("Startup completed", StringComparison.Ordinal) >= 0
                || message.IndexOf("Baseline completed", StringComparison.Ordinal) >= 0;
        }
    }
}
