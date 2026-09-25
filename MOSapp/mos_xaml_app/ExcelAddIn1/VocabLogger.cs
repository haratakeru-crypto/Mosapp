using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelAddIn1
{
    /// <summary>
    /// 単語帳モード用イベントを %TEMP%\mos_excel_vocab_events.txt に追記する。
    /// </summary>
    public static class VocabLogger
    {
        static readonly object LockObject = new object();
        static readonly string EventFilePath = Path.Combine(Path.GetTempPath(), "mos_excel_vocab_events.txt");
        static readonly string ModeFilePath = Path.Combine(Path.GetTempPath(), "mos_excel_vocab_mode.txt");

        public static bool IsVocabModeEnabled()
        {
            try
            {
                if (!File.Exists(ModeFilePath)) return false;
                string t = File.ReadAllText(ModeFilePath).Trim();
                return t == "1" || t.Equals("true", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        public static void LogKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            if (!IsVocabModeEnabled()) return;

            try
            {
                lock (LockObject)
                {
                    string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [Vocab] Key={key.Trim()}{Environment.NewLine}";
                    File.AppendAllText(EventFilePath, line, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[VocabLogger] " + ex.Message);
            }
        }

        public static void LogFormulaIfAny(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) return;
            string trimmed = formula.Trim();
            if (!trimmed.StartsWith("=")) return;

            var m = Regex.Match(trimmed, @"^=\s*([A-Za-z][A-Za-z0-9\.]*)");
            if (!m.Success) return;
            LogKey("Formula:" + m.Groups[1].Value.ToUpperInvariant());
        }
    }
}
