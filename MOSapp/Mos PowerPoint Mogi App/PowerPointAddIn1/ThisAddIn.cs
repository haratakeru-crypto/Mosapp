using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private static readonly string CurrentTaskFilePath = Path.Combine(Path.GetTempPath(), "mos_ppt_current_task.txt");

        private Timer _grayscalePollTimer;
        private bool _lastBlackAndWhite;
        private Timer _audio8_4PollTimer;
        private bool _task8_4Logged;
        private Timer _layout10_7PollTimer;
        private bool _task10_7Logged;
        private Timer _printOptionsPollTimer;
        private string _lastPrintPresFullName;
        private int _lastPrintOutputType = -1;
        private int _lastPrintCopies = -1;
        private int _lastPrintCollate = -1;
        private bool _printOptionsInitialized;
        private bool _task5_1PrintLogged;
        private bool _task11_7PrintLogged;

        private Timer _kiosk7_4PollTimer;
        private bool _task7_4KioskLogged;
        private Timer _task1_2To1_4PollTimer;
        private bool _task1_2Logged;
        private bool _task1_3Logged;
        private bool _task1_4Logged;
        private List<int> _task1_4PrevSlideIds = new List<int>();
        private string _task1_4PrevPresentationKey;

        private Timer _taskFilePollTimer;
        private int _currentTaskProjectId = -1;
        private int _currentTaskTaskId = -1;
        private int _currentTaskAttemptNo = 1;

        private const float PositionTolerancePt = 0.5f;

        /// <summary>現在タスク（VSTO が読み取った ProjectId, TaskId）。ログ記録時に使用。</summary>
        internal static int CurrentTaskProjectId { get; private set; } = -1;
        /// <summary>現在タスクの TaskId。</summary>
        internal static int CurrentTaskTaskId { get; private set; } = -1;

        private int _currentTaskExemptFlags = 0;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[PowerPointAddIn1] Add-in started. Log file: " + Logger.GetLogFilePath());

            _lastBlackAndWhite = false;
            _grayscalePollTimer = new Timer();
            _grayscalePollTimer.Interval = 500;
            _grayscalePollTimer.Tick += GrayscalePollTimer_Tick;
            _grayscalePollTimer.Start();

            _task5_1PrintLogged = false;
            _task11_7PrintLogged = false;
            _printOptionsPollTimer = new Timer();
            _printOptionsPollTimer.Interval = 1000;
            _printOptionsPollTimer.Tick += PrintOptionsPollTimer_Tick;
            _printOptionsPollTimer.Start();

            _task8_4Logged = false;
            _audio8_4PollTimer = new Timer();
            _audio8_4PollTimer.Interval = 1000;
            _audio8_4PollTimer.Tick += Audio8_4PollTimer_Tick;
            _audio8_4PollTimer.Start();

            _task10_7Logged = false;
            _layout10_7PollTimer = new Timer();
            _layout10_7PollTimer.Interval = 1500;
            _layout10_7PollTimer.Tick += Layout10_7PollTimer_Tick;
            _layout10_7PollTimer.Start();

            _task7_4KioskLogged = false;
            _kiosk7_4PollTimer = new Timer();
            _kiosk7_4PollTimer.Interval = 1000;
            _kiosk7_4PollTimer.Tick += Kiosk7_4PollTimer_Tick;
            _kiosk7_4PollTimer.Start();

            _task1_2Logged = false;
            _task1_3Logged = false;
            _task1_4Logged = false;
            _task1_4PrevSlideIds.Clear();
            _task1_4PrevPresentationKey = null;
            _task1_2To1_4PollTimer = new Timer();
            _task1_2To1_4PollTimer.Interval = 700;
            _task1_2To1_4PollTimer.Tick += Task1_2To1_4PollTimer_Tick;
            _task1_2To1_4PollTimer.Start();

            _taskFilePollTimer = new Timer();
            _taskFilePollTimer.Interval = 500;
            _taskFilePollTimer.Tick += TaskFilePollTimer_Tick;
            _taskFilePollTimer.Start();
        }

        private void TaskFilePollTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!File.Exists(CurrentTaskFilePath))
                {
                    // 最終タスク（11-7）でレビュー遷移時に current_task が消えるケースでも、
                    // 離脱直前の印刷設定を1回だけ再評価して証跡を確定する。
                    if (_currentTaskProjectId == 11 && _currentTaskTaskId == 7)
                    {
                        TryLogTask11_7PrintOnTaskBoundary();
                    }
                    _currentTaskProjectId = -1;
                    _currentTaskTaskId = -1;
                    _currentTaskAttemptNo = 1;
                    return;
                }
                string line = null;
                try
                {
                    line = File.ReadAllText(CurrentTaskFilePath).Trim();
                    if (string.IsNullOrEmpty(line)) return;
                }
                catch { return; }
                var parts = line.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 2) return;
                if (!int.TryParse(parts[0].Trim(), out int projectId) || !int.TryParse(parts[1].Trim(), out int taskId))
                    return;

                bool forceSnapshot = !File.Exists(SnapshotFilePath);

                if (projectId == _currentTaskProjectId && taskId == _currentTaskTaskId && !forceSnapshot)
                    return;

                // --- タスク切り替え時の処理 ---
                // 5-1 はポーリング取りこぼし対策として、タスク離脱直前に印刷設定を即時再評価して証跡を確定する。
                if (_currentTaskProjectId == 5 && _currentTaskTaskId == 1)
                {
                    TryLogTask5_1PrintOnTaskBoundary();
                }
                // 11-7 も同様に、タスク離脱直前の即時再評価で証跡を確定する。
                if (_currentTaskProjectId == 11 && _currentTaskTaskId == 7)
                {
                    TryLogTask11_7PrintOnTaskBoundary();
                }

                // 新しいタスクを開始する前に、直前のタスクの破壊的操作チェックを行う
                // ※ プロジェクトIDが変わる場合は、比較対象のプレゼンテーションが異なるためスキップする
                if (_currentTaskProjectId != -1 && !forceSnapshot && projectId == _currentTaskProjectId)
                {
                    CheckAndLogDestructiveOperations(_currentTaskProjectId, _currentTaskTaskId, _currentTaskExemptFlags);
                }

                _currentTaskProjectId = projectId;
                _currentTaskTaskId = taskId;
                _currentTaskExemptFlags = parts.Length >= 3 ? int.Parse(parts[2].Trim()) : 0;
                int attemptNo = 1;
                if (parts.Length >= 4)
                {
                    int.TryParse(parts[3].Trim(), out attemptNo);
                    if (attemptNo < 1) attemptNo = 1;
                }
                _currentTaskAttemptNo = attemptNo;
                if (!(projectId == 1 && taskId == 2)) _task1_2Logged = false;
                if (!(projectId == 1 && taskId == 3)) _task1_3Logged = false;
                if (!(projectId == 1 && taskId == 4))
                {
                    _task1_4Logged = false;
                    _task1_4PrevSlideIds.Clear();
                    _task1_4PrevPresentationKey = null;
                }

                CurrentTaskProjectId = projectId;
                CurrentTaskTaskId = taskId;
                Logger.SetCurrentTaskContext(projectId, taskId, attemptNo);
                Logger.LogTaskStart(projectId, taskId, attemptNo);
                TakeUnifiedSnapshot();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[TaskFilePoll] " + ex.Message);
            }
        }

        private static readonly string SnapshotFilePath = Path.Combine(Path.GetTempPath(), "mos_ppt_snapshot.txt");
        private static readonly string DestructiveLogPath = Path.Combine(Path.GetTempPath(), "mos_ppt_destructive_errors.log");

        private void TakeUnifiedSnapshot()
        {
            try
            {
                var status = CaptureCurrentStatus(CurrentTaskProjectId, CurrentTaskTaskId);
                if (status == null) return;
                SaveSnapshot(status);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[TakeSnapshot] " + ex.Message);
            }
        }

        private SnapshotData CaptureCurrentStatus(int pid, int tid)
        {
            try
            {
                if (Application == null) return null;
                PowerPoint.Presentation pres = null;
                try { pres = Application.ActivePresentation; } catch { }
                if (pres == null) return null;

                var data = new SnapshotData { ProjectId = pid, TaskId = tid };
                var slides = pres.Slides;
                if (slides != null)
                {
                    data.SlidesCount = slides.Count;
                    for (int i = 1; i <= data.SlidesCount; i++)
                    {
                        PowerPoint.Slide slide = slides[i];
                        data.SlideNames.Add(slide.Name);
                        
                        PowerPoint.Shapes shapes = slide.Shapes;
                        data.ShapesCounts[i] = shapes.Count;
                        
                        long slideTextLength = 0;
                        for (int j = 1; j <= shapes.Count; j++)
                        {
                            PowerPoint.Shape shape = shapes[j];
                            try
                            {
                                // Text Check (TextFrame2 priority)
                                try {
                                    dynamic tf2 = shape.TextFrame2;
                                    if (tf2 != null && (int)tf2.HasText == -1) slideTextLength += tf2.TextRange.Length;
                                    else if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                        slideTextLength += shape.TextFrame.TextRange.Length;
                                } catch { }

                                // Position Check
                                if (IsImageOrPlaceholder(shape))
                                {
                                    string key = $"{i}_{shape.Id}";
                                    data.ShapePositions[key] = Tuple.Create(shape.Left, shape.Top, shape.Width, shape.Height);
                                }
                            }
                            catch { }
                            finally { Marshal.ReleaseComObject(shape); }
                        }
                        
                        int animCount = 0;
                        try { animCount = slide.TimeLine.MainSequence.Count; } catch { }
                        data.AnimationCounts[i] = animCount;
                        data.TotalTextLength += slideTextLength;
                        data.SlideTextLengths[i] = slideTextLength;

                        Marshal.ReleaseComObject(shapes);
                        Marshal.ReleaseComObject(slide);
                    }
                    Marshal.ReleaseComObject(slides);
                }
                return data;
            }
            catch { return null; }
        }

        private List<string> CompareSnapshots(SnapshotData start, SnapshotData current, int exemptFlagsInt)
        {
            var errors = new List<string>();
            var flags = (PPValidationExemptFlags)exemptFlagsInt;

            if (!flags.HasFlag(PPValidationExemptFlags.SlidesCount))
            {
                if (current.SlidesCount != start.SlidesCount) errors.Add("SlidesCount changed");
            }

            if (!flags.HasFlag(PPValidationExemptFlags.ShapesCount))
            {
                foreach (var kvp in start.ShapesCounts)
                {
                    if (current.ShapesCounts.ContainsKey(kvp.Key) && current.ShapesCounts[kvp.Key] != kvp.Value)
                        errors.Add($"ShapesCount on Slide {kvp.Key} changed");
                }
            }
            else
            {
                // 免除されているが、厳密なデルタチェックを適用
                foreach (var kvp in start.ShapesCounts)
                {
                    int allowedDelta = GetAllowedShapesCountDelta(start.ProjectId, start.TaskId, kvp.Key);
                    if (allowedDelta != int.MaxValue)
                    {
                        if (current.ShapesCounts.ContainsKey(kvp.Key))
                        {
                            int actualDelta = current.ShapesCounts[kvp.Key] - kvp.Value;
                            if (!IsAllowedShapesCountDelta(start.ProjectId, start.TaskId, kvp.Key, allowedDelta, actualDelta))
                            {
                                errors.Add(FormatDestructiveShapesCountMessage(kvp.Key, start.ProjectId, start.TaskId, allowedDelta, actualDelta));
                            }
                        }
                    }
                }
            }

            if (!flags.HasFlag(PPValidationExemptFlags.TextLength))
            {
                if (current.TotalTextLength != start.TotalTextLength) errors.Add("TotalTextLength changed");
            }
            else
            {
                foreach (var kvp in start.SlideTextLengths)
                {
                    int allowedDelta = GetAllowedTextLengthDelta(start.ProjectId, start.TaskId, kvp.Key);
                    if (allowedDelta != int.MaxValue)
                    {
                        if (current.SlideTextLengths.ContainsKey(kvp.Key))
                        {
                            long actualDelta = current.SlideTextLengths[kvp.Key] - kvp.Value;
                            if (!IsAllowedTextLengthDelta(start.ProjectId, start.TaskId, kvp.Key, allowedDelta, actualDelta))
                            {
                                errors.Add(FormatDestructiveTextLengthMessage(kvp.Key, start.ProjectId, start.TaskId, allowedDelta, actualDelta));
                            }
                        }
                    }
                }
            }

            // 図形座標・サイズの比較
            bool exemptFullShapePosition = flags.HasFlag(PPValidationExemptFlags.ShapePosition);
            bool onlyNewShapesExempt = IsShapePositionExemptForNewShapesOnly(start.ProjectId, start.TaskId);
            int allowedExistingChangesCount = GetAllowedExistingShapePositionChangeCount(start.ProjectId, start.TaskId);

            if (!exemptFullShapePosition || onlyNewShapesExempt || allowedExistingChangesCount >= 0)
            {
                int changedExistingShapesCount = 0;
                foreach (var kvp in start.ShapePositions)
                {
                    if (current.ShapePositions.ContainsKey(kvp.Key))
                    {
                        var cPos = current.ShapePositions[kvp.Key];
                        var sPos = kvp.Value;
                        if (Math.Abs(cPos.Item1 - sPos.Item1) > PositionTolerancePt ||
                            Math.Abs(cPos.Item2 - sPos.Item2) > PositionTolerancePt ||
                            Math.Abs(cPos.Item3 - sPos.Item3) > PositionTolerancePt ||
                            Math.Abs(cPos.Item4 - sPos.Item4) > PositionTolerancePt)
                        {
                            if (exemptFullShapePosition)
                            {
                                if (onlyNewShapesExempt)
                                {
                                    errors.Add($"不正な図形変更: 指示外の既存図形(ID:{kvp.Key})の位置・サイズが変更されています。");
                                }
                                else if (allowedExistingChangesCount >= 0)
                                {
                                    changedExistingShapesCount++;
                                    if (changedExistingShapesCount > allowedExistingChangesCount)
                                    {
                                        errors.Add($"上限超過の図形変更: 許可された数以上の既存図形(ID:{kvp.Key})が変更されています。");
                                    }
                                }
                            }
                            else
                            {
                                errors.Add($"Shape position/size changed on Slide {kvp.Key.Split('_')[0]} (ID:{kvp.Key})");
                            }
                        }
                    }
                }
            }
            return errors;
        }

        private static bool IsAllowedShapesCountDelta(int projectId, int taskId, int slideIndex, int allowedDelta, int actualDelta)
        {
            if (allowedDelta == int.MaxValue) return true;

            // 6-3: Slide 1 only: allow 0 or +1. Disallow deletions (<0) and bulk additions (>1).
            if (projectId == 6 && taskId == 3 && slideIndex == 1)
            {
                return actualDelta == 0 || actualDelta == 1;
            }

            return actualDelta == allowedDelta;
        }

        private static bool IsAllowedTextLengthDelta(int projectId, int taskId, int slideIndex, int allowedDelta, long actualDelta)
        {
            if (allowedDelta == int.MaxValue) return true;

            // 9-6 slide 1: allow 0 (already replaced at snapshot) or -57 (expected URL→お問い合わせ).
            if (projectId == 9 && taskId == 6 && slideIndex == 1)
            {
                return actualDelta == 0 || actualDelta == -57;
            }

            return actualDelta == allowedDelta;
        }

        private static string FormatDestructiveShapesCountMessage(int slideIndex, int projectId, int taskId, int allowedDelta, int actualDelta)
        {
            if (projectId == 6 && taskId == 3 && slideIndex == 1)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +1、実際の変化: {actualDelta}）";
            }
            return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（期待される変化数: {allowedDelta}、実際: {actualDelta}）";
        }

        private static string FormatDestructiveTextLengthMessage(int slideIndex, int projectId, int taskId, int allowedDelta, long actualDelta)
        {
            if (projectId == 9 && taskId == 6 && slideIndex == 1)
            {
                return $"不正なテキスト変更: スライド {slideIndex} で指示外のテキスト変更が検知されました（許容: 文字数の変化は 0 または -57、実際の変化: {actualDelta}）";
            }
            return $"不正なテキスト変更: スライド {slideIndex} で指示外のテキスト変更が検知されました（期待される文字数変化: {allowedDelta}、実際: {actualDelta}）";
        }

        private bool IsShapePositionExemptForNewShapesOnly(int projectId, int taskId)
        {
            if (projectId == 3 && (taskId == 1 || taskId == 3 || taskId == 4)) return true; // 3-1, 3-3, 3-4
            if (projectId == 4 && taskId == 6) return true; // 4-6
            if (projectId == 5 && (taskId == 3 || taskId == 5)) return true; // 5-3, 5-5
            if (projectId == 6 && taskId == 3) return true; // 6-3
            if (projectId == 9 && taskId == 1) return true; // 9-1
            if (projectId == 10 && taskId == 7) return true; // 10-7
            return false;
        }

        private int GetAllowedExistingShapePositionChangeCount(int projectId, int taskId)
        {
            if (projectId == 4 && taskId == 4) return 1; // 4-4
            if (projectId == 4 && taskId == 5) return 1; // 4-5
            if (projectId == 5 && taskId == 4) return 1; // 5-4
            if (projectId == 6 && taskId == 4) return 1; // 6-4
            if (projectId == 9 && taskId == 1) return -1; // デフォルトへ (deltaで制御)
            if (projectId == 9 && taskId == 6) return 1; // 9-6
            if (projectId == 11 && taskId == 6) return 1; // 11-6
            return -1;
        }

        private int GetAllowedShapesCountDelta(int projectId, int taskId, int slideIndex)
        {
            if (projectId == 3 && taskId == 1) return slideIndex == 5 ? 0 : 0; // 3-1
            if (projectId == 3 && taskId == 3) return slideIndex == 6 ? 0 : 0; // 3-3
            if (projectId == 3 && taskId == 4) return slideIndex == 1 ? 2 : 0; // 3-4
            if (projectId == 5 && taskId == 3) return 0;                       // 5-3
            if (projectId == 5 && taskId == 5) return slideIndex == 3 ? -2 : 0; // 5-5
            if (projectId == 6 && taskId == 3) return slideIndex == 1 ? 1 : 0; // 6-3
            if (projectId == 9 && taskId == 1) return slideIndex == 2 ? 0 : 0; // 9-1

            return int.MaxValue;
        }

        private int GetAllowedTextLengthDelta(int projectId, int taskId, int slideIndex)
        {
            // 1-7: 吹き出しへのテキスト入力 (スライド1に「教育者必見」の5文字が追加される)
            if (projectId == 1 && taskId == 7) return slideIndex == 1 ? 5 : 0;
            // 9-6: URLを「お問い合わせ」に変更 (スライド1の63文字のURLが6文字の「お問い合わせ」に置き換わるため -57文字)
            if (projectId == 9 && taskId == 6) return slideIndex == 1 ? -57 : 0;

            // 変換、削除、インポートなど文字数が可変なものはチェックを省略
            return int.MaxValue;
        }

        private void SaveSnapshot(SnapshotData data)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"TaskId:{data.ProjectId},{data.TaskId}");
                sb.AppendLine($"SlidesCount:{data.SlidesCount}");
                sb.AppendLine($"TotalTextLength:{data.TotalTextLength}");
                
                var shapeCountsStr = string.Join("|", data.ShapesCounts.Select(x => $"{x.Key}:{x.Value}"));
                sb.AppendLine($"ShapesCounts:{shapeCountsStr}");

                var textLengthsStr = string.Join("|", data.SlideTextLengths.Select(x => $"{x.Key}:{x.Value}"));
                sb.AppendLine($"SlideTextLengths:{textLengthsStr}");

                var posList = data.ShapePositions.Select(x => $"{x.Key}:{x.Value.Item1},{x.Value.Item2},{x.Value.Item3},{x.Value.Item4}");
                sb.AppendLine($"ShapePositions:{string.Join("|", posList)}");

                File.WriteAllText(SnapshotFilePath, sb.ToString());
            }
            catch { }
        }

        private SnapshotData LoadSnapshot()
        {
            if (!File.Exists(SnapshotFilePath)) return null;
            var data = new SnapshotData();
            try
            {
                var lines = File.ReadAllLines(SnapshotFilePath);
                foreach (var line in lines)
                {
                    var idx = line.IndexOf(':');
                    if (idx < 0) continue;
                    var key = line.Substring(0, idx);
                    var val = line.Substring(idx + 1);
                    switch (key)
                    {
                        case "TaskId":
                            var ids = val.Split(',');
                            data.ProjectId = int.Parse(ids[0]);
                            data.TaskId = int.Parse(ids[1]);
                            break;
                        case "SlidesCount": data.SlidesCount = int.Parse(val); break;
                        case "TotalTextLength": data.TotalTextLength = long.Parse(val); break;
                        case "ShapesCounts":
                            foreach (var part in val.Split('|')) {
                                var kv = part.Split(':');
                                if (kv.Length == 2) data.ShapesCounts[int.Parse(kv[0])] = int.Parse(kv[1]);
                            }
                            break;
                        case "SlideTextLengths":
                            foreach (var part in val.Split('|')) {
                                var kv = part.Split(':');
                                if (kv.Length == 2) data.SlideTextLengths[int.Parse(kv[0])] = long.Parse(kv[1]);
                            }
                            break;
                        case "ShapePositions":
                            foreach (var part in val.Split('|')) {
                                var kv = part.Split(':');
                                if (kv.Length == 2) {
                                    var coords = kv[1].Split(',');
                                    data.ShapePositions[kv[0]] = Tuple.Create(float.Parse(coords[0]), float.Parse(coords[1]), float.Parse(coords[2]), float.Parse(coords[3]));
                                }
                            }
                            break;
                    }
                }
                return data;
            }
            catch { return null; }
        }

        private class SnapshotData
        {
            public int ProjectId; public int TaskId; public int SlidesCount;
            public List<string> SlideNames = new List<string>();
            public Dictionary<int, int> ShapesCounts = new Dictionary<int, int>();
            public long TotalTextLength;
            public Dictionary<int, long> SlideTextLengths = new Dictionary<int, long>();
            public Dictionary<int, int> AnimationCounts = new Dictionary<int, int>();
            public Dictionary<string, Tuple<float, float, float, float>> ShapePositions = new Dictionary<string, Tuple<float, float, float, float>>();
        }

        [Flags]
        private enum PPValidationExemptFlags
        {
            None = 0, ShapesCount = 1, TextLength = 2, SlidesCount = 4, AnimationRemoved = 8, ShapePosition = 16, All = 31
        }

        private void CheckAndLogDestructiveOperations(int projectId, int taskId, int exemptFlagsInt)
        {
            try
            {
                // 現在のスナップショット（開始時のデータ）をロード
                var startSnapshot = LoadSnapshot();
                if (startSnapshot == null || startSnapshot.ProjectId != projectId || startSnapshot.TaskId != taskId) return;

                // 現在のリアルタイムな状態を取得
                var currentStatus = CaptureCurrentStatus(projectId, taskId);
                if (currentStatus == null) return;

                // 比較
                List<string> errors = CompareSnapshots(startSnapshot, currentStatus, exemptFlagsInt);
                if (errors.Count > 0)
                {
                    // ログに記録
                    string errorMsg = string.Join(" | ", errors);
                    File.AppendAllText(DestructiveLogPath, $"{projectId},{taskId},{_currentTaskAttemptNo}:{errorMsg}{Environment.NewLine}");
                    System.Diagnostics.Debug.WriteLine($"[DestructiveCheck] Task {projectId}-{taskId} FAILED: {errorMsg}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[CheckAndLog] " + ex.Message);
            }
        }

        private static bool IsImageOrPlaceholder(PowerPoint.Shape sh)
        {
            try
            {
                int t = (int)sh.Type;
                if (t == (int)Office.MsoShapeType.msoPicture) return true;
                if (t == (int)Office.MsoShapeType.msoPlaceholder) return true;
                try
                {
                    var pf = sh.PlaceholderFormat;
                    if (pf != null) { Marshal.ReleaseComObject(pf); return true; }
                }
                catch { }
                return false;
            }
            catch { return false; }
        }


        private void Layout10_7PollTimer_Tick(object sender, EventArgs e)
        {
            if (_task10_7Logged) return;
            if (!IsCurrentTask(10, 7)) return;
            try
            {
                if (Application == null || Application.Presentations == null) return;
                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;
                    PowerPoint.Master master = null;
                    try
                    {
                        master = pres.SlideMaster;
                        if (master == null) return;
                        PowerPoint.CustomLayouts layouts = null;
                        try
                        {
                            layouts = master.CustomLayouts;
                            if (layouts == null) return;
                            for (int i = 1; i <= layouts.Count; i++)
                            {
                                PowerPoint.CustomLayout cl = null;
                                try
                                {
                                    cl = layouts[i];
                                    if (cl == null) continue;
                                    string name = null;
                                    try { name = cl.Name ?? ""; } catch { continue; }
                                    if (name.IndexOf("画像付きスライド", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        Logger.LogTask10_7LayoutDuplicate();
                                        _task10_7Logged = true;
                                        return;
                                    }
                                }
                                finally { if (cl != null) try { Marshal.ReleaseComObject(cl); } catch { } }
                            }
                        }
                        finally { if (layouts != null) try { Marshal.ReleaseComObject(layouts); } catch { } }
                    }
                    finally { if (master != null) try { Marshal.ReleaseComObject(master); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch { }
        }

        private void PrintOptionsPollTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                bool isTask5_1 = IsCurrentTask(5, 1);
                bool isTask11_7 = IsCurrentTask(11, 7);
                if (!isTask5_1 && !isTask11_7) return;
                if (Application == null || Application.Presentations == null) return;
                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;
                    string fullName = null;
                    try { fullName = pres.FullName ?? ""; } catch { return; }
                    if (string.IsNullOrEmpty(fullName)) fullName = pres.Name ?? "";

                    if (_lastPrintPresFullName != null && fullName != _lastPrintPresFullName)
                    {
                        _lastPrintPresFullName = null;
                        _printOptionsInitialized = false;
                        _lastPrintOutputType = -1;
                        _lastPrintCopies = -1;
                        _lastPrintCollate = -1;
                    _task7_4KioskLogged = false;
                    }

                    PowerPoint.PrintOptions po = null;
                    try
                    {
                        po = pres.PrintOptions;
                        if (po == null) return;
                        int outputType = (int)po.OutputType;
                        int copies = po.NumberOfCopies;
                        int collateInt = Convert.ToInt32(po.Collate);
                        bool collate = (collateInt == (int)Office.MsoTriState.msoTrue);

                        if (!_printOptionsInitialized)
                        {
                            _lastPrintPresFullName = fullName;
                            _lastPrintOutputType = outputType;
                            _lastPrintCopies = copies;
                            _lastPrintCollate = collateInt;
                            _printOptionsInitialized = true;
                            return;
                        }

                        bool changed = (_lastPrintOutputType != outputType || _lastPrintCopies != copies || _lastPrintCollate != collateInt);
                        _lastPrintOutputType = outputType;
                        _lastPrintCopies = copies;
                        _lastPrintCollate = collateInt;

                        if (changed)
                        {
                            if (isTask5_1 && !_task5_1PrintLogged &&
                                outputType == (int)PowerPoint.PpPrintOutputType.ppPrintOutputThreeSlideHandouts &&
                                copies == 4 && collate)
                            {
                                Logger.LogTask5_1Print();
                                _task5_1PrintLogged = true;
                            }
                            if (isTask11_7 && !_task11_7PrintLogged &&
                                outputType == (int)PowerPoint.PpPrintOutputType.ppPrintOutputNotesPages &&
                                copies == 3 && collate)
                            {
                                Logger.LogTask11_7Print();
                                _task11_7PrintLogged = true;
                            }
                        }
                    }
                    finally { if (po != null) try { Marshal.ReleaseComObject(po); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch { }
        }

        /// <summary>
        /// 5-1 から離脱する直前に印刷設定を即時確認し、条件一致なら証跡ログを確定する。
        /// ポーリング間隔中の取りこぼしを補完するための境界処理。
        /// </summary>
        private void TryLogTask5_1PrintOnTaskBoundary()
        {
            try
            {
                if (_task5_1PrintLogged) return;
                if (Application == null || Application.Presentations == null) return;

                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;

                    PowerPoint.PrintOptions po = null;
                    try
                    {
                        po = pres.PrintOptions;
                        if (po == null) return;

                        int outputType = (int)po.OutputType;
                        int copies = po.NumberOfCopies;
                        bool collate = (Convert.ToInt32(po.Collate) == (int)Office.MsoTriState.msoTrue);

                        if (outputType == (int)PowerPoint.PpPrintOutputType.ppPrintOutputThreeSlideHandouts
                            && copies == 4
                            && collate)
                        {
                            Logger.LogTask5_1Print();
                            _task5_1PrintLogged = true;
                        }
                    }
                    finally { if (po != null) try { Marshal.ReleaseComObject(po); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch { }
        }

        /// <summary>
        /// 11-7 から離脱する直前に印刷設定を即時確認し、条件一致なら証跡ログを確定する。
        /// 最終タスクでレビュー遷移時に current_task が消える経路の取りこぼしも補完する。
        /// </summary>
        private void TryLogTask11_7PrintOnTaskBoundary()
        {
            try
            {
                if (_task11_7PrintLogged) return;
                if (Application == null || Application.Presentations == null) return;

                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;

                    PowerPoint.PrintOptions po = null;
                    try
                    {
                        po = pres.PrintOptions;
                        if (po == null) return;

                        int outputType = (int)po.OutputType;
                        int copies = po.NumberOfCopies;
                        bool collate = (Convert.ToInt32(po.Collate) == (int)Office.MsoTriState.msoTrue);

                        if (outputType == (int)PowerPoint.PpPrintOutputType.ppPrintOutputNotesPages
                            && copies == 3
                            && collate)
                        {
                            Logger.LogTask11_7Print();
                            _task11_7PrintLogged = true;
                        }
                    }
                    finally { if (po != null) try { Marshal.ReleaseComObject(po); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch { }
        }

        private void Audio8_4PollTimer_Tick(object sender, EventArgs e)
        {
            if (_task8_4Logged) return;
            if (!IsCurrentTask(8, 4)) return;
            try
            {
                if (Application == null || Application.Presentations == null) return;
                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;
                    PowerPoint.Slide slide = null;
                    try
                    {
                        PowerPoint.Slides slides = pres.Slides;
                        if (slides == null || slides.Count < 1) return;
                        slide = slides[1];
                        if (slide == null) return;
                        PowerPoint.Shapes shapes = slide.Shapes;
                        if (shapes == null) return;
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PowerPoint.Shape sh = null;
                            try
                            {
                                sh = shapes[i];
                                try
                                {
                                    if (sh.MediaType != PowerPoint.PpMediaType.ppMediaTypeSound) continue;
                                }
                                catch { continue; }
                                PowerPoint.MediaFormat mf = null;
                                try
                                {
                                    mf = sh.MediaFormat;
                                    if (mf == null) continue;
                                    float fadeIn = (float)mf.FadeInDuration;
                                    if (Math.Abs(fadeIn - 4000f) < 500f)
                                    {
                                        Logger.LogTask8_4Audio();
                                        _task8_4Logged = true;
                                        return;
                                    }
                                }
                                finally { if (mf != null) try { Marshal.ReleaseComObject(mf); } catch { } }
                            }
                            finally { if (sh != null) try { Marshal.ReleaseComObject(sh); } catch { } }
                        }
                    }
                    finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch { }
        }

        private void GrayscalePollTimer_Tick(object sender, EventArgs e)
        {
            if (!IsCurrentTask(10, 4))
            {
                _lastBlackAndWhite = false;
                return;
            }
            try
            {
                if (Application == null) return;
                dynamic window = Application.ActiveWindow;
                if (window == null) return;

                bool current = false;
                try
                {
                    // COM は MsoTriState を整数で返すため、enum 比較ではなく数値で判定する
                    current = (Convert.ToInt32(window.BlackAndWhite) == (int)Office.MsoTriState.msoTrue);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[GrayscalePoll] BlackAndWhite get failed: " + ex.Message);
                    return;
                }

                if (current && !_lastBlackAndWhite)
                {
                    Logger.LogTask10_4Grayscale();
                }
                _lastBlackAndWhite = current;
            }
            catch
            {
                // アドインが落ちないように握りつぶす
            }
        }

        private void Kiosk7_4PollTimer_Tick(object sender, EventArgs e)
        {
            if (_task7_4KioskLogged) return;
            if (!IsCurrentTask(7, 4)) return;
            try
            {
                if (Application == null || Application.Presentations == null) return;
                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;
                    PowerPoint.SlideShowSettings ss = null;
                    try
                    {
                        ss = pres.SlideShowSettings;
                        if (ss == null) return;
                        if (ss.ShowType == PowerPoint.PpSlideShowType.ppShowTypeKiosk)
                        {
                            Logger.LogTask7_4Kiosk();
                            _task7_4KioskLogged = true;
                        }
                    }
                    finally { if (ss != null) try { Marshal.ReleaseComObject(ss); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch { }
        }

        private void Task1_2To1_4PollTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (Application == null || Application.Presentations == null) return;
                if (!IsCurrentTask(1, 2) && !IsCurrentTask(1, 3) && !IsCurrentTask(1, 4)) return;

                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;

                    if (IsCurrentTask(1, 2) && !_task1_2Logged)
                    {
                        PowerPoint.Slides slides = null;
                        PowerPoint.Slide slide2 = null;
                        PowerPoint.Slide slide3 = null;
                        PowerPoint.CustomLayout l2 = null;
                        PowerPoint.CustomLayout l3 = null;
                        try
                        {
                            slides = pres.Slides;
                            if (slides != null && slides.Count >= 3)
                            {
                                slide2 = slides[2];
                                slide3 = slides[3];
                                if (slide2 != null && slide3 != null)
                                {
                                    l2 = slide2.CustomLayout;
                                    l3 = slide3.CustomLayout;
                                    string n2 = l2?.Name ?? "";
                                    string n3 = l3?.Name ?? "";
                                    if (!string.IsNullOrEmpty(n2) && string.Equals(n2, n3, StringComparison.OrdinalIgnoreCase))
                                    {
                                        Logger.LogTask1_2Duplicate();
                                        _task1_2Logged = true;
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (l2 != null) try { Marshal.ReleaseComObject(l2); } catch { }
                            if (l3 != null) try { Marshal.ReleaseComObject(l3); } catch { }
                            if (slide2 != null) try { Marshal.ReleaseComObject(slide2); } catch { }
                            if (slide3 != null) try { Marshal.ReleaseComObject(slide3); } catch { }
                            if (slides != null) try { Marshal.ReleaseComObject(slides); } catch { }
                        }
                    }

                    if (IsCurrentTask(1, 3) && !_task1_3Logged)
                    {
                        PowerPoint.Slides slides = null;
                        PowerPoint.Slide slide3 = null;
                        try
                        {
                            slides = pres.Slides;
                            if (slides != null && slides.Count >= 3)
                            {
                                slide3 = slides[3];
                                if (slide3 != null && slide3.SlideShowTransition.Hidden == Office.MsoTriState.msoTrue)
                                {
                                    Logger.LogTask1_3Hide();
                                    _task1_3Logged = true;
                                }
                            }
                        }
                        finally
                        {
                            if (slide3 != null) try { Marshal.ReleaseComObject(slide3); } catch { }
                            if (slides != null) try { Marshal.ReleaseComObject(slides); } catch { }
                        }
                    }

                    if (IsCurrentTask(1, 4) && !_task1_4Logged)
                    {
                        string key = null;
                        try { key = (pres.FullName ?? pres.Name ?? "").Trim(); } catch { }
                        PowerPoint.Slides slides = null;
                        try
                        {
                            slides = pres.Slides;
                            if (slides == null) return;
                            var currentIds = new List<int>();
                            for (int i = 1; i <= slides.Count; i++)
                            {
                                PowerPoint.Slide s = null;
                                try
                                {
                                    s = slides[i];
                                    if (s != null) currentIds.Add(s.SlideID);
                                }
                                finally { if (s != null) try { Marshal.ReleaseComObject(s); } catch { } }
                            }

                            if (string.IsNullOrEmpty(_task1_4PrevPresentationKey) || !string.Equals(_task1_4PrevPresentationKey, key, StringComparison.OrdinalIgnoreCase))
                            {
                                _task1_4PrevPresentationKey = key;
                                _task1_4PrevSlideIds = new List<int>(currentIds);
                                return;
                            }

                            if (_task1_4PrevSlideIds.Count >= 3 && currentIds.Count < _task1_4PrevSlideIds.Count)
                            {
                                int deletedId = _task1_4PrevSlideIds[2];
                                if (!currentIds.Contains(deletedId))
                                {
                                    Logger.LogTask1_4DeleteThirdSlide();
                                    _task1_4Logged = true;
                                }
                            }
                            _task1_4PrevSlideIds = new List<int>(currentIds);
                        }
                        finally { if (slides != null) try { Marshal.ReleaseComObject(slides); } catch { } }
                    }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_taskFilePollTimer != null)
            {
                _taskFilePollTimer.Stop();
                _taskFilePollTimer.Dispose();
                _taskFilePollTimer = null;
            }
            if (_kiosk7_4PollTimer != null)
            {
                _kiosk7_4PollTimer.Stop();
                _kiosk7_4PollTimer.Dispose();
                _kiosk7_4PollTimer = null;
            }
            if (_task1_2To1_4PollTimer != null)
            {
                _task1_2To1_4PollTimer.Stop();
                _task1_2To1_4PollTimer.Dispose();
                _task1_2To1_4PollTimer = null;
            }
            if (_printOptionsPollTimer != null)
            {
                _printOptionsPollTimer.Stop();
                _printOptionsPollTimer.Dispose();
                _printOptionsPollTimer = null;
            }
            if (_audio8_4PollTimer != null)
            {
                _audio8_4PollTimer.Stop();
                _audio8_4PollTimer.Dispose();
                _audio8_4PollTimer = null;
            }
            if (_layout10_7PollTimer != null)
            {
                _layout10_7PollTimer.Stop();
                _layout10_7PollTimer.Dispose();
                _layout10_7PollTimer = null;
            }
            if (_grayscalePollTimer != null)
            {
                _grayscalePollTimer.Stop();
                _grayscalePollTimer.Dispose();
                _grayscalePollTimer = null;
            }
            System.Diagnostics.Debug.WriteLine("[PowerPointAddIn1] Add-in shutdown");
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            System.Diagnostics.Debug.WriteLine("[ThisAddIn] CreateRibbonExtensibilityObject called");
            return new Ribbon();
        }

        private static bool IsCurrentTask(int projectId, int taskId)
        {
            return CurrentTaskProjectId == projectId && CurrentTaskTaskId == taskId;
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
