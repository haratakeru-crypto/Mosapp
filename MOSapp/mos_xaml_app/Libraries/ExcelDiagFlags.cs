using System;
using System.IO;

namespace Libraries
{
    /// <summary>
    /// 開発用ログの出し分け。Debug ビルドでは常に詳細を書く。
    /// Release／製品では %TEMP%\mos_excel_diag.on があるときだけ書く。
    /// 採点（mos_excel_log.txt の [Op]）と起動待ち（Startup / Baseline completed）は常に残す。
    /// </summary>
    public static class ExcelDiagFlags
    {
        public const string FlagFileName = "mos_excel_diag.on";

        public static string FlagPath => Path.Combine(Path.GetTempPath(), FlagFileName);

        /// <summary>
        /// 詳細ログを出してよいか。Release のフラグ判定は数秒キャッシュするので、
        /// 製品で調査するときはファイル作成後に少し待つか再起動する。
        /// </summary>
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

        /// <summary>
        /// VSTO 診断ファイルへ書いてよいか。起動待ちに使う行は常に許可する。
        /// </summary>
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
