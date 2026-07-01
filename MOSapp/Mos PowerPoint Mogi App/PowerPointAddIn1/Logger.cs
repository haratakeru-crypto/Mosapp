using System;
using System.IO;
using System.Text;

namespace PowerPointAddIn1
{
    /// <summary>
    /// PowerPoint 操作ログを記録するクラス。
    /// mos_ppt_log.txt に全般を記録するほか、4-3/5-1/10-4/11-7 の採点根拠は mos_ppt_task_evidence.txt にも追記する（単体リセットでメインログが消されても採点可能にする）。
    /// </summary>
    public static class Logger
    {
        private static readonly object _lockObject = new object();
        private static readonly string _logFileName = "mos_ppt_log.txt";
        private static readonly string _taskEvidenceFileName = "mos_ppt_task_evidence.txt";
        private static string _logFilePath;
        private static string _taskEvidenceFilePath;
        private static int _currentProjectId = -1;
        private static int _currentTaskId = -1;
        private static int _currentAttemptNo = 1;

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

        /// <summary>採点用証跡（PPLogReader.GetTaskEvidenceLogPath と同一パス）</summary>
        private static string TaskEvidenceFilePath
        {
            get
            {
                if (string.IsNullOrEmpty(_taskEvidenceFilePath))
                    _taskEvidenceFilePath = Path.Combine(Path.GetTempPath(), _taskEvidenceFileName);
                return _taskEvidenceFilePath;
            }
        }

        /// <summary>
        /// コマンド実行をログに記録
        /// </summary>
        public static void LogCommand(string commandId)
        {
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string logEntry = $"[{timestamp}] [{commandId}] Executed";
                    AppendToFile(LogFilePath, logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
            }
        }

        /// <summary>
        /// 10-4 グレースケール操作をログに記録（採点で参照する形式）
        /// </summary>
        public static void LogTask10_4Grayscale()
        {
            // 他の証跡タスク（5-1, 11-7）と同様、Task コンテキスト付きの統一フォーマットで記録する。
            LogTaskTag("Task10-4", "Grayscale");
        }

        /// <summary>P8-5: 表示グレースケール操作をログに記録。</summary>
        public static void LogTask8_5Grayscale()
        {
            LogTaskTag("Task8-5", "Grayscale");
        }

        /// <summary>P8-3: スライドサイズ16:9設定をログに記録（P8-4で上書きされるため一括採点用）。</summary>
        public static void LogTask8_3SlideSize16x9()
        {
            LogTaskTag("Task8-3", "SlideSize16x9");
        }

        /// <summary>P2-1: スプリット＋ワイプアウト（横）の画面切り替えを記録（P2-4で上書きされるため一括採点用）。</summary>
        public static void LogTask2_1SplitHorizontalOut()
        {
            LogTaskTag("Task2-1", "SplitHorizontalOut");
        }

        /// <summary>P2-2: 全スライド画面切り替え継続時間3秒を記録（P2-4で上書きされるため一括採点用）。</summary>
        public static void LogTask2_2TransitionDuration3Sec()
        {
            LogTaskTag("Task2-2", "TransitionDuration3Sec");
        }

        /// <summary>P2-3: スライド3〜5に「切り替え」を記録（P2-4で上書きされるため一括採点用）。</summary>
        public static void LogTask2_3SwitchRight()
        {
            LogTaskTag("Task2-3", "SwitchRight");
        }

        /// <summary>5-1: 配布資料3スライド・部単位4部で印刷したことを記録。</summary>
        public static void LogTask5_1Print()
        {
            LogTaskTag("Task5-1", "Print");
        }

        /// <summary>11-7: ノート・3部・部単位で印刷したことを記録。</summary>
        public static void LogTask11_7Print()
        {
            LogTaskTag("Task11-7", "Print");
        }

        /// <summary>P6-5: アウトライン・6部・部単位で印刷設定したことを記録。</summary>
        public static void LogTask6_5Print()
        {
            LogTaskTag("Task6-5", "Print");
        }

        /// <summary>P6-6: ノート・3部・ページ単位（Collate OFF）で印刷設定したことを記録。</summary>
        public static void LogTask6_6Print()
        {
            LogTaskTag("Task6-6", "Print");
        }

        /// <summary>P6-7: グレースケール配布資料3スライド/頁・4部で印刷設定したことを記録。</summary>
        public static void LogTask6_7Print()
        {
            LogTaskTag("Task6-7", "Print");
        }

        /// <summary>4-3: スライド1画像に光彩18pt・アクセント6を適用したことを記録。</summary>
        public static void LogTask4_3Glow()
        {
            LogTaskTag("Task4-3", "Glow18Accent6");
        }

        /// <summary>1-2: スライド2の複製状態（2,3枚目同レイアウト）を検出したことを記録。</summary>
        public static void LogTask1_2Duplicate()
        {
            LogTaskTag("Task1-2", "Duplicate");
        }

        /// <summary>1-3: スライド3を非表示にしたことを記録。</summary>
        public static void LogTask1_3Hide()
        {
            LogTaskTag("Task1-3", "HideSlide3");
        }

        /// <summary>1-4: （開始時の）3枚目スライドを削除したことを記録。</summary>
        public static void LogTask1_4DeleteThirdSlide()
        {
            LogTaskTag("Task1-4", "DeleteThirdSlide");
        }

        /// <summary>1-8: サマリーズームスライドが挿入されたことを記録。</summary>
        public static void LogTask1_8SummaryZoom()
        {
            LogTaskTag("Task1-8", "SummaryZoom");
        }

        /// <summary>8-4: オーディオ再生設定（フェードイン4秒等）を記録。Legacy 参照用。</summary>
        public static void LogTask8_4Audio()
        {
            LogTaskTag("Task8-4", "Audio");
        }

        /// <summary>P9-3: スライド切替後も再生を記録（旧8-4）。</summary>
        public static void LogTask9_3PlayAcrossSlides()
        {
            LogTaskTag("Task9-3", "PlayAcrossSlides");
        }

        /// <summary>P9-3: フェードアウト3秒を記録（旧8-4）。</summary>
        public static void LogTask9_3FadeOut3000()
        {
            LogTaskTag("Task9-3", "FadeOut3000");
        }

        /// <summary>7-2: スライドの再利用を記録。</summary>
        public static void LogTask7_2ReuseSlides()
        {
            LogTaskTag("Task7-2", "ReuseSlides");
        }

        /// <summary>7-3: アウトラインから挿入を記録。</summary>
        public static void LogTask7_3InsertFromOutline()
        {
            LogTaskTag("Task7-3", "InsertFromOutline");
        }

        /// <summary>7-4: スライドショーを自動プレゼンテーション（Kiosk）に設定したことを記録。</summary>
        public static void LogTask7_4Kiosk()
        {
            LogTaskTag("Task7-4", "Kiosk");
        }

        /// <summary>P6-3: スライドショーを自動プレゼンテーション（Kiosk）に設定したことを記録。</summary>
        public static void LogTask6_3Kiosk()
        {
            LogTaskTag("Task6-3", "Kiosk");
        }

        /// <summary>10-1: ドキュメント検査実行を記録。</summary>
        public static void LogTask10_1DocumentInspector()
        {
            LogTaskTag("Task10-1", "DocumentInspector");
        }

        /// <summary>10-7: レイアウト複製を記録。</summary>
        public static void LogTask10_7LayoutDuplicate()
        {
            LogTaskTag("Task10-7", "LayoutDuplicate");
        }

        /// <summary>
        /// 現在タスク開始を記録（試験アプリがタスクを切り替えたとき）。形式: [timestamp] [TaskStart] P-T
        /// </summary>
        public static void LogTaskStart(int projectId, int taskId, int attemptNo = 1)
        {
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    if (attemptNo < 1) attemptNo = 1;
                    string logEntry = $"[{timestamp}] [TaskStart] {projectId}-{taskId}-{attemptNo}";
                    AppendToFile(LogFilePath, logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
            }
        }

        public static void SetCurrentTaskContext(int projectId, int taskId, int attemptNo)
        {
            lock (_lockObject)
            {
                _currentProjectId = projectId;
                _currentTaskId = taskId;
                _currentAttemptNo = attemptNo < 1 ? 1 : attemptNo;
            }
        }

        /// <summary>
        /// 汎用操作を記録。現在タスクが指定されていれば [Task P-T] をプレフィックスする。形式: [timestamp] [Task P-T] [Op] Type Detail
        /// </summary>
        public static void LogOperation(string operationType, string detail, int? projectId = null, int? taskId = null)
        {
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string taskPrefix = (projectId.HasValue && taskId.HasValue) ? $"[Task {projectId.Value}-{taskId.Value}] " : "";
                    string logEntry = $"[{timestamp}] {taskPrefix}[Op] {operationType} {detail}";
                    AppendToFile(LogFilePath, logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
            }
        }

        private static void LogTaskTag(string taskTag, string identifier)
        {
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string contextPrefix = (_currentProjectId > 0 && _currentTaskId > 0)
                        ? $"[Task {_currentProjectId}-{_currentTaskId}-{_currentAttemptNo}] "
                        : "";
                    string logEntry = $"[{timestamp}] {contextPrefix}[{taskTag}] {identifier}";
                    AppendToFile(LogFilePath, logEntry);
                    // 採点がログ依存のタスクは証跡にも同一行を残す（単体プロジェクトリセットでメインログ削除後も採点可能）
                    if (string.Equals(taskTag, "Task1-2", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task1-3", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task1-4", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task1-8", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task4-3", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task5-1", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task6-5", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task6-6", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task6-7", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task11-7", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task10-4", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task8-5", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task8-3", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task9-3", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task2-1", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task2-2", StringComparison.Ordinal)
                        || string.Equals(taskTag, "Task2-3", StringComparison.Ordinal))
                    {
                        AppendToFile(TaskEvidenceFilePath, logEntry);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
            }
        }

        private static void AppendToFile(string filePath, string logEntry)
        {
            using (var fileStream = new FileStream(
                filePath,
                FileMode.Append,
                FileAccess.Write,
                FileShare.ReadWrite | FileShare.Delete))
            using (var writer = new StreamWriter(fileStream, Encoding.UTF8))
            {
                writer.WriteLine(logEntry);
            }
        }

        /// <summary>
        /// メイン操作ログのみクリア。採点用証跡（mos_ppt_task_evidence.txt）は消さない。
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
        /// ログファイルのパスを取得（採点側・デバックタブ用）
        /// </summary>
        public static string GetLogFilePath()
        {
            return LogFilePath;
        }
    }
}
