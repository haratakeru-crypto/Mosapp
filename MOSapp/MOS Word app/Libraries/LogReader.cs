using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Libraries
{
    /// <summary>
    /// VSTOアドインで生成されたログファイルを読み込むユーティリティクラス
    /// </summary>
    public static class LogReader
    {
        /// <summary>
        /// ログファイルのパスを取得
        /// </summary>
        public static string GetLogFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_word_log.txt");
        }

        /// <summary>
        /// VSTOアドインのログファイルをクリアする。
        /// 「すべてリセット」「リセット」実行時に呼び出し、採点で参照するログを初期化する。
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
                System.Diagnostics.Debug.WriteLine($"[LogReader] Error clearing log: {ex.Message}");
            }
        }

        /// <summary>
        /// ログファイルからすべてのログエントリを読み込む
        /// </summary>
        /// <returns>ログエントリのリスト（時系列順）</returns>
        public static List<LogEntry> ReadAllLogEntries()
        {
            var entries = new List<LogEntry>();
            string logFilePath = GetLogFilePath();

            if (!File.Exists(logFilePath))
            {
                return entries;
            }

            try
            {
                string[] lines = File.ReadAllLines(logFilePath);
                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    LogEntry entry = ParseLogLine(line);
                    if (entry != null)
                    {
                        entries.Add(entry);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LogReader] Error reading log file: {ex.Message}");
            }

            return entries;
        }

        /// <summary>
        /// 指定したコマンドIDが実行された回数をカウント
        /// </summary>
        /// <param name="commandId">コマンドID（例: "ShowAll", "Cut"）</param>
        /// <returns>実行回数</returns>
        public static int CountCommandExecution(string commandId)
        {
            var entries = ReadAllLogEntries();
            return entries.Count(e => e.CommandId.Equals(commandId, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 指定したコマンドIDが実行されたかをチェック
        /// </summary>
        /// <param name="commandId">コマンドID</param>
        /// <returns>実行された場合true</returns>
        public static bool HasCommandExecuted(string commandId)
        {
            return CountCommandExecution(commandId) > 0;
        }

        /// <summary>
        /// いずれかのコマンドIDがログに記録されているか（例: 4-3 の ReviewResolveComment / ReviewDeleteComment）
        /// </summary>
        public static bool HasAnyCommandExecuted(params string[] commandIds)
        {
            if (commandIds == null || commandIds.Length == 0)
                return false;
            foreach (string id in commandIds)
            {
                if (string.IsNullOrWhiteSpace(id))
                    continue;
                if (HasCommandExecuted(id.Trim()))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// 指定したコマンドIDの実行が、指定回数以上あるかをチェック
        /// </summary>
        /// <param name="commandId">コマンドID</param>
        /// <param name="minCount">最小実行回数</param>
        /// <returns>指定回数以上実行された場合true</returns>
        public static bool HasCommandExecutedAtLeast(string commandId, int minCount)
        {
            return CountCommandExecution(commandId) >= minCount;
        }

        /// <summary>
        /// ログ行をパースしてLogEntryオブジェクトに変換
        /// 形式: [yyyy-MM-dd HH:mm:ss] [CommandID] Executed
        /// </summary>
        private static LogEntry ParseLogLine(string line)
        {
            try
            {
                // [yyyy-MM-dd HH:mm:ss] [CommandID] Executed の形式をパース
                int firstBracketEnd = line.IndexOf(']');
                if (firstBracketEnd < 0)
                    return null;

                string timestampStr = line.Substring(1, firstBracketEnd - 1);
                if (!DateTime.TryParse(timestampStr, out DateTime timestamp))
                    return null;

                int secondBracketStart = line.IndexOf('[', firstBracketEnd + 1);
                if (secondBracketStart < 0)
                    return null;

                int secondBracketEnd = line.IndexOf(']', secondBracketStart + 1);
                if (secondBracketEnd < 0)
                    return null;

                string commandId = line.Substring(secondBracketStart + 1, secondBracketEnd - secondBracketStart - 1);

                return new LogEntry
                {
                    Timestamp = timestamp,
                    CommandId = commandId
                };
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// ログエントリを表すクラス
        /// </summary>
        public class LogEntry
        {
            public DateTime Timestamp { get; set; }
            public string CommandId { get; set; }

            public override string ToString()
            {
                return $"[{Timestamp:yyyy-MM-dd HH:mm:ss}] [{CommandId}] Executed";
            }
        }
    }
}

