using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace New_MOSWordVSTOAddIn
{
    /// <summary>
    /// 段階診断: %TEMP%\mos_word_diag_empty_addin.txt
    ///   0  heartbeat のみ
    ///   1  DocumentChange のみ
    ///   2  全ポーリング（Fast+ShowAllPoll+Destructive）、リボン無し
    ///   3  リボン（コマンドフックあり）、ポーリング無し
    ///   4  リボン＋全ポーリング（DocumentChange 無し）
    ///   5  Fast タイマーのみ（300ms ShowAll/heartbeat/flush）
    ///   6  ShowAllPoll のみ（1200ms）
    ///   7  Destructive のみ（500ms current_task）
    ///   8  リボン殻のみ（IRibbonExtensibility あるが idMso コマンド無し）
    /// フラグ無し → 通常（DocumentChange 無し・リボン有効）
    /// </summary>
    internal static class DiagMode
    {
        private static readonly string FlagPath = Path.Combine(Path.GetTempPath(), "mos_word_diag_empty_addin.txt");
        private static readonly string LogPath = Path.Combine(Path.GetTempPath(), "mos_word_diag.log");
        private static int? _stageCached;
        private static int _documentChangeSeq;
        private static DateTime _lastDocumentChangeUtc = DateTime.MinValue;

        public const int StageNormal = -1;
        public const int StageHeartbeatOnly = 0;
        public const int StageEvents = 1;
        public const int StagePollsOnly = 2;
        public const int StageRibbonOnly = 3;
        public const int StageRibbonAndPolls = 4;
        public const int StageFastPollOnly = 5;
        public const int StageSlowPollOnly = 6;
        public const int StageDestructiveOnly = 7;
        public const int StageRibbonShellOnly = 8;

        public const int DetailSlowMs = 50;

        public static string GetFlagPath() => FlagPath;
        public static string GetLogFilePath() => LogPath;

        public static bool IsDiagMode() => GetStage() >= StageHeartbeatOnly;

        public static bool TraceDocumentChange() => IsDiagMode() && EnableDocumentChange();

        public static bool EnableRibbon()
        {
            int s = GetStage();
            return s == StageNormal
                || s == StageRibbonOnly
                || s == StageRibbonAndPolls
                || s == StageRibbonShellOnly;
        }

        /// <summary>idMso コマンドフック付きフルリボンか。stage 8 は殻のみ。</summary>
        public static bool EnableRibbonCommands()
        {
            int s = GetStage();
            return s == StageNormal || s == StageRibbonOnly || s == StageRibbonAndPolls;
        }

        public static bool EnableDocumentChange() => GetStage() == StageEvents;

        public static bool EnableDocumentBeforeSave()
        {
            int s = GetStage();
            return s == StageNormal || s == StageEvents || s == StageRibbonAndPolls;
        }

        public static bool EnableFastPoll()
        {
            int s = GetStage();
            return s == StageNormal || s == StagePollsOnly || s == StageRibbonAndPolls || s == StageFastPollOnly;
        }

        public static bool EnableSlowPoll()
        {
            int s = GetStage();
            return s == StageNormal || s == StagePollsOnly || s == StageRibbonAndPolls || s == StageSlowPollOnly;
        }

        public static bool EnableDestructivePoll()
        {
            int s = GetStage();
            return s == StageNormal || s == StagePollsOnly || s == StageRibbonAndPolls || s == StageDestructiveOnly;
        }

        public static bool EnablePollingTimers() =>
            EnableFastPoll() || EnableSlowPoll() || EnableDestructivePoll();

        public static int GetStage()
        {
            if (_stageCached.HasValue)
                return _stageCached.Value;

            try
            {
                if (!File.Exists(FlagPath))
                {
                    _stageCached = StageNormal;
                    return _stageCached.Value;
                }

                string text = "";
                try { text = File.ReadAllText(FlagPath).Trim(); }
                catch { /* ignore */ }

                if (string.IsNullOrEmpty(text))
                    _stageCached = StageHeartbeatOnly;
                else if (int.TryParse(text.Split(new[] { '\r', '\n', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)[0], out int n))
                    _stageCached = Math.Max(0, Math.Min(8, n));
                else if (text.StartsWith("event", StringComparison.OrdinalIgnoreCase))
                    _stageCached = StageEvents;
                else if (text.StartsWith("fast", StringComparison.OrdinalIgnoreCase))
                    _stageCached = StageFastPollOnly;
                else if (text.StartsWith("slow", StringComparison.OrdinalIgnoreCase))
                    _stageCached = StageSlowPollOnly;
                else if (text.StartsWith("destr", StringComparison.OrdinalIgnoreCase))
                    _stageCached = StageDestructiveOnly;
                else if (text.StartsWith("shell", StringComparison.OrdinalIgnoreCase))
                    _stageCached = StageRibbonShellOnly;
                else if (text.StartsWith("poll", StringComparison.OrdinalIgnoreCase)
                    || text.StartsWith("timer", StringComparison.OrdinalIgnoreCase))
                    _stageCached = StagePollsOnly;
                else if (text.StartsWith("ribbon", StringComparison.OrdinalIgnoreCase))
                    _stageCached = StageRibbonOnly;
                else if (text.StartsWith("both", StringComparison.OrdinalIgnoreCase)
                    || text.StartsWith("full", StringComparison.OrdinalIgnoreCase))
                    _stageCached = StageRibbonAndPolls;
                else if (text.StartsWith("empty", StringComparison.OrdinalIgnoreCase)
                    || text.StartsWith("heart", StringComparison.OrdinalIgnoreCase))
                    _stageCached = StageHeartbeatOnly;
                else
                    _stageCached = StageHeartbeatOnly;
            }
            catch
            {
                _stageCached = StageNormal;
            }

            return _stageCached.Value;
        }

        public static string StageName(int stage)
        {
            switch (stage)
            {
                case StageHeartbeatOnly: return "0=heartbeat-only";
                case StageEvents: return "1=DocumentChange-only";
                case StagePollsOnly: return "2=all-polls (no ribbon)";
                case StageRibbonOnly: return "3=ribbon+commands (no polls)";
                case StageRibbonAndPolls: return "4=ribbon+polls";
                case StageFastPollOnly: return "5=fast-poll-only (300ms)";
                case StageSlowPollOnly: return "6=showAllPoll-only (1200ms)";
                case StageDestructiveOnly: return "7=destructive-only (500ms)";
                case StageRibbonShellOnly: return "8=ribbon-shell (no idMso hooks)";
                default: return "normal";
            }
        }

        public static void Write(string message)
        {
            if (string.IsNullOrEmpty(message))
                return;
            try
            {
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + message + Environment.NewLine;
                File.AppendAllText(LogPath, line, Encoding.UTF8);
                System.Diagnostics.Debug.WriteLine("[DiagMode] " + message);
            }
            catch
            {
                // ignore
            }
        }

        public static int BeginDocumentChangeTrace(out long sinceLastMs)
        {
            DateTime now = DateTime.UtcNow;
            sinceLastMs = _lastDocumentChangeUtc == DateTime.MinValue
                ? -1
                : (long)(now - _lastDocumentChangeUtc).TotalMilliseconds;
            _lastDocumentChangeUtc = now;
            _documentChangeSeq++;
            return _documentChangeSeq;
        }

        public static void Measure(string name, Action action, bool alwaysLog = false)
        {
            if (action == null)
                return;
            if (!TraceDocumentChange())
            {
                action();
                return;
            }

            var sw = Stopwatch.StartNew();
            try { action(); }
            finally
            {
                sw.Stop();
                if (alwaysLog || sw.ElapsedMilliseconds >= DetailSlowMs)
                    Write("  " + name + " " + sw.ElapsedMilliseconds + "ms");
            }
        }

        public static T Measure<T>(string name, Func<T> func, bool alwaysLog = false)
        {
            if (func == null)
                return default(T);
            if (!TraceDocumentChange())
                return func();

            var sw = Stopwatch.StartNew();
            try { return func(); }
            finally
            {
                sw.Stop();
                if (alwaysLog || sw.ElapsedMilliseconds >= DetailSlowMs)
                    Write("  " + name + " " + sw.ElapsedMilliseconds + "ms");
            }
        }

        public static void LogStartupBanner()
        {
            int stage = GetStage();
            if (stage < StageHeartbeatOnly)
            {
                Write("===== normal add-in (diag flag absent; DocumentChange off; Ribbon on) =====");
                return;
            }

            Write("===== ADD-IN DIAG stage " + StageName(stage) + " =====");
            Write("flag=" + FlagPath);
            Write("ribbon=" + EnableRibbon()
                + " ribbonCommands=" + EnableRibbonCommands()
                + " documentChange=" + EnableDocumentChange()
                + " beforeSave=" + EnableDocumentBeforeSave()
                + " fastPoll=" + EnableFastPoll()
                + " slowPoll=" + EnableSlowPoll()
                + " destructive=" + EnableDestructivePoll());
            Write("scoring may be incomplete — diagnostic use only");
        }
    }
}
