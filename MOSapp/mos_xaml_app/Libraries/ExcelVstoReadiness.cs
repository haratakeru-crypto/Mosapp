using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Libraries
{
    /// <summary>
    /// VSTO 診断ログを増分で読み、対象 Excel の Startup completed だけを共有判定する。
    /// </summary>
    public static class ExcelVstoReadiness
    {
        public const string StartupCompletedMarker = "Startup completed";
        public const string BaselineCompletedMarker = "Baseline completed";

        private static readonly Regex PidRegex = new Regex(@"\[PID:(\d+)\]", RegexOptions.Compiled);
        private static readonly Regex TokenRegex = new Regex(@"token=([A-Za-z0-9]+)", RegexOptions.Compiled);
        private static readonly object Sync = new object();
        private static readonly HashSet<int> ReadyProcessIds = new HashSet<int>();
        private static readonly Dictionary<int, HashSet<string>> ReadyTokens =
            new Dictionary<int, HashSet<string>>();
        private static readonly string DiagnosticPath = ExcelLogReader.GetDiagnosticLogPath();
        private static readonly string OpenTokenPath = Path.Combine(Path.GetTempPath(), "mos_excel_open_token.txt");

        private static long _readOffset;
        private static string _pendingLine = "";

        public static string CreateOpenToken()
        {
            string token = Guid.NewGuid().ToString("N");
            try
            {
                File.WriteAllText(OpenTokenPath, token, new UTF8Encoding(false));
            }
            catch
            {
                /* 書けない場合はトークン無しで従来の Startup completed 待ちに戻す */
            }
            return token;
        }

        public static bool IsStartupCompleted(int processId)
        {
            if (processId <= 0)
                return false;

            Refresh();
            lock (Sync)
            {
                return ReadyProcessIds.Contains(processId);
            }
        }

        /// <summary>
        /// 今回のオープンで作った基準の完了だけを見る。過去の Startup completed では閉じない。
        /// </summary>
        public static bool IsOpenReady(int processId, string openToken)
        {
            if (processId <= 0)
                return false;
            if (string.IsNullOrEmpty(openToken))
                return IsStartupCompleted(processId);

            Refresh();
            lock (Sync)
            {
                HashSet<string> tokens;
                return ReadyTokens.TryGetValue(processId, out tokens) && tokens.Contains(openToken);
            }
        }

        public static bool WaitForStartup(int processId, int timeoutMs)
        {
            if (processId <= 0)
                return false;

            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (IsStartupCompleted(processId))
                    return true;
                Thread.Sleep(200);
            }

            return IsStartupCompleted(processId);
        }

        public static void RecordHostEvent(string message)
        {
            if (string.IsNullOrEmpty(message))
                return;

            try
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
                File.AppendAllText(
                    DiagnosticPath,
                    "[" + timestamp + "] [Host] " + message + Environment.NewLine,
                    new UTF8Encoding(false));
            }
            catch
            {
                /* 診断ログが書けなくても起動は続ける */
            }
        }

        private static void Refresh()
        {
            lock (Sync)
            {
                try
                {
                    if (!File.Exists(DiagnosticPath))
                        return;

                    using (var stream = new FileStream(
                        DiagnosticPath,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite))
                    {
                        if (stream.Length < _readOffset)
                        {
                            _readOffset = 0;
                            _pendingLine = "";
                            ReadyProcessIds.Clear();
                            ReadyTokens.Clear();
                        }

                        if (stream.Length == _readOffset)
                            return;

                        stream.Seek(_readOffset, SeekOrigin.Begin);
                        int remaining = (int)Math.Min(int.MaxValue, stream.Length - _readOffset);
                        var buffer = new byte[remaining];
                        int read = stream.Read(buffer, 0, buffer.Length);
                        _readOffset += read;
                        if (read <= 0)
                            return;

                        string chunk = _pendingLine + Encoding.UTF8.GetString(buffer, 0, read);
                        int consumed = 0;
                        for (int i = 0; i < chunk.Length; i++)
                        {
                            if (chunk[i] != '\n')
                                continue;
                            int length = i - consumed;
                            if (length > 0 && chunk[i - 1] == '\r')
                                length--;
                            ConsumeLine(chunk.Substring(consumed, length));
                            consumed = i + 1;
                        }

                        _pendingLine = consumed < chunk.Length
                            ? chunk.Substring(consumed)
                            : "";
                    }
                }
                catch
                {
                    /* 読めない間は未完了のまま待つ */
                }
            }
        }

        private static void ConsumeLine(string line)
        {
            int processId;
            string token;
            bool startup;
            if (!TryParseReadyLine(line, out processId, out token, out startup))
                return;
            if (startup)
                ReadyProcessIds.Add(processId);
            if (string.IsNullOrEmpty(token))
                return;

            HashSet<string> tokens;
            if (!ReadyTokens.TryGetValue(processId, out tokens))
            {
                tokens = new HashSet<string>(StringComparer.Ordinal);
                ReadyTokens[processId] = tokens;
            }
            tokens.Add(token);
        }

        public static bool TryParseStartupCompletedLine(string line, out int processId)
        {
            string token;
            bool startup;
            return TryParseReadyLine(line, out processId, out token, out startup) && startup;
        }

        public static bool TryParseReadyLine(string line, out int processId, out string token, out bool startupCompleted)
        {
            processId = 0;
            token = "";
            startupCompleted = false;
            if (string.IsNullOrEmpty(line) || line.IndexOf("[Host]", StringComparison.Ordinal) >= 0)
                return false;

            startupCompleted = line.IndexOf(StartupCompletedMarker, StringComparison.Ordinal) >= 0;
            bool baseline = line.IndexOf(BaselineCompletedMarker, StringComparison.Ordinal) >= 0;
            if (!startupCompleted && !baseline)
                return false;

            Match pidMatch = PidRegex.Match(line);
            if (!pidMatch.Success)
                return false;
            if (!int.TryParse(pidMatch.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out processId)
                || processId <= 0)
                return false;

            Match tokenMatch = TokenRegex.Match(line);
            if (tokenMatch.Success)
                token = tokenMatch.Groups[1].Value;
            return true;
        }
    }
}
