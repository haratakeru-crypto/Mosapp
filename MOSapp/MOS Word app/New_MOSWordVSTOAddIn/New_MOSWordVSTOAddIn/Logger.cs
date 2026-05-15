using System;
using System.IO;
using System.Text;

namespace New_MOSWordVSTOAddIn
{
    /// <summary>
    /// Word操作ログを記録するクラス。
    /// リボンコマンド（編集記号の表示 ShowAll など）の実行のみをログに残す。
    /// 文字の入力・削除等のログは行わない。
    /// </summary>
    public static class Logger
    {
        private static readonly object _lockObject = new object();
        private static readonly string _logFileName = "mos_word_log.txt";
        private static string _logFilePath;

        /// <summary>
        /// ログファイルのパスを取得
        /// </summary>
        private static string LogFilePath
        {
            get
            {
                if (string.IsNullOrEmpty(_logFilePath))
                {
                    _logFilePath = Path.Combine(Path.GetTempPath(), _logFileName);
                }
                return _logFilePath;
            }
        }

        /// <summary>
        /// コマンド実行をログに記録
        /// </summary>
        /// <param name="commandId">コマンドID</param>
        public static void LogCommand(string commandId)
        {
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string logEntry = $"[{timestamp}] [{commandId}] Executed";

                    // ログファイルに追記
                    // 外部プロセス（MOS Word アプリ側）がログファイルを削除できるように
                    // FileShare.ReadWrite | FileShare.Delete で開く
                    using (var fileStream = new FileStream(
                        LogFilePath,
                        FileMode.Append,
                        FileAccess.Write,
                        FileShare.ReadWrite | FileShare.Delete))
                    using (var writer = new StreamWriter(fileStream, Encoding.UTF8))
                    {
                        writer.WriteLine(logEntry);
                    }
                }
            }
            catch (Exception ex)
            {
                // ログ書き込み失敗時もWordの動作に影響を与えない
                // デバッグビルド時のみデバッグ出力
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
            }
        }

        /// <summary>
        /// ログファイルをクリア
        /// </summary>
        public static void ClearLog()
        {
            try
            {
                lock (_lockObject)
                {
                    if (File.Exists(LogFilePath))
                    {
                        File.Delete(LogFilePath);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error clearing log: {ex.Message}");
            }
        }

        /// <summary>
        /// ログファイルのパスを取得（外部からの読み込み用）
        /// </summary>
        public static string GetLogFilePath()
        {
            return LogFilePath;
        }
    }
}




