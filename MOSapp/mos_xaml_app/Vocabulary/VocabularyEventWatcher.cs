using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>
    /// VSTO が %TEMP%\mos_excel_vocab_events.txt に追記するイベントを監視する。
    /// 行形式: [Vocab] Key=...
    /// </summary>
    public sealed class VocabularyEventWatcher : IDisposable
    {
        public static readonly string EventFilePath =
            Path.Combine(Path.GetTempPath(), "mos_excel_vocab_events.txt");

        readonly FileSystemWatcher _watcher;
        long _readPosition;
        readonly object _sync = new object();

        public event Action<string> EventReceived;

        public VocabularyEventWatcher()
        {
            EnsureFile();
            _readPosition = new FileInfo(EventFilePath).Length;

            _watcher = new FileSystemWatcher(Path.GetDirectoryName(EventFilePath) ?? Path.GetTempPath())
            {
                Filter = Path.GetFileName(EventFilePath),
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size
            };
            _watcher.Changed += (_, __) => Drain();
            _watcher.EnableRaisingEvents = true;
        }

        public static void ClearEvents()
        {
            try
            {
                File.WriteAllText(EventFilePath, string.Empty, Encoding.UTF8);
            }
            catch { }
        }

        public void Drain()
        {
            lock (_sync)
            {
                try
                {
                    if (!File.Exists(EventFilePath)) return;
                    using (var fs = new FileStream(EventFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        if (fs.Length < _readPosition)
                            _readPosition = 0;
                        fs.Seek(_readPosition, SeekOrigin.Begin);
                        using (var reader = new StreamReader(fs, Encoding.UTF8))
                        {
                            string line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                string key = ParseKey(line);
                                if (!string.IsNullOrEmpty(key))
                                    EventReceived?.Invoke(key);
                            }
                            _readPosition = fs.Position;
                        }
                    }
                }
                catch { }
            }
        }

        static string ParseKey(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return null;
            var m = Regex.Match(line, @"\[Vocab\]\s*Key=(.+)$", RegexOptions.IgnoreCase);
            if (!m.Success) return null;
            return m.Groups[1].Value.Trim();
        }

        static void EnsureFile()
        {
            try
            {
                if (!File.Exists(EventFilePath))
                    File.WriteAllText(EventFilePath, string.Empty, Encoding.UTF8);
            }
            catch { }
        }

        public void Dispose()
        {
            try
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
            }
            catch { }
        }
    }
}
