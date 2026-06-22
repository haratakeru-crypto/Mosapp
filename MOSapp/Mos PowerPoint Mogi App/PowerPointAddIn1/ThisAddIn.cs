using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Libraries;
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
        private Timer _glow4_3PollTimer;
        private bool _task4_3GlowLogged;
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
        private bool _task1_8Logged;
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

            _task4_3GlowLogged = false;
            _glow4_3PollTimer = new Timer();
            _glow4_3PollTimer.Interval = 1000;
            _glow4_3PollTimer.Tick += Glow4_3PollTimer_Tick;
            _glow4_3PollTimer.Start();

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

            _task1_8Logged = false;
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
                // 11-7 はポーリング取りこぼし対策として、タスク離脱直前に印刷設定を即時再評価して証跡を確定する。
                if (_currentTaskProjectId == 11 && _currentTaskTaskId == 7)
                {
                    TryLogTask11_7PrintOnTaskBoundary();
                }
                // 4-3 は光彩が 4-4 で外れるため、離脱直前に COM/OpenXML で証跡を確定する。
                if (_currentTaskProjectId == 4 && _currentTaskTaskId == 3)
                {
                    TryLogTask4_3GlowOnTaskBoundary();
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
                if (!(projectId == 1 && taskId == 8)) _task1_8Logged = false;
                if (!(projectId == 4 && taskId == 3)) _task4_3GlowLogged = false;

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

            if (UsesSlideIndexMapping(start.ProjectId, start.TaskId))
            {
                if (!IsSlidesCountValidForTask(start.ProjectId, start.TaskId, start.SlidesCount, current.SlidesCount))
                    errors.Add("SlidesCount changed");
            }
            else if (!flags.HasFlag(PPValidationExemptFlags.SlidesCount))
            {
                if (current.SlidesCount != start.SlidesCount) errors.Add("SlidesCount changed");
            }

            if (!flags.HasFlag(PPValidationExemptFlags.ShapesCount))
            {
                foreach (var kvp in start.ShapesCounts)
                {
                    int currentSlide = MapSnapshotSlideToCurrent(start.ProjectId, start.TaskId, kvp.Key, start.SlidesCount, current.SlidesCount);
                    if (current.ShapesCounts.ContainsKey(currentSlide) && current.ShapesCounts[currentSlide] != kvp.Value)
                        errors.Add($"ShapesCount on Slide {currentSlide} changed");
                }
            }
            else
            {
                foreach (var kvp in start.ShapesCounts)
                {
                    int snapshotSlide = kvp.Key;
                    int currentSlide = MapSnapshotSlideToCurrent(start.ProjectId, start.TaskId, snapshotSlide, start.SlidesCount, current.SlidesCount);
                    int allowedDelta = GetAllowedShapesCountDelta(start.ProjectId, start.TaskId, currentSlide);
                    if (allowedDelta != int.MaxValue)
                    {
                        if (current.ShapesCounts.ContainsKey(currentSlide))
                        {
                            int actualDelta = current.ShapesCounts[currentSlide] - kvp.Value;
                            if (!IsAllowedShapesCountDelta(start.ProjectId, start.TaskId, currentSlide, allowedDelta, actualDelta))
                            {
                                errors.Add(FormatDestructiveShapesCountMessage(currentSlide, start.ProjectId, start.TaskId, allowedDelta, actualDelta));
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
                    int snapshotSlide = kvp.Key;
                    int currentSlide = MapSnapshotSlideToCurrent(start.ProjectId, start.TaskId, snapshotSlide, start.SlidesCount, current.SlidesCount);
                    int allowedDelta = GetAllowedTextLengthDelta(start.ProjectId, start.TaskId, currentSlide);
                    if (allowedDelta != int.MaxValue)
                    {
                        if (current.SlideTextLengths.ContainsKey(currentSlide))
                        {
                            long actualDelta = current.SlideTextLengths[currentSlide] - kvp.Value;
                            if (!IsAllowedTextLengthDelta(start.ProjectId, start.TaskId, currentSlide, allowedDelta, actualDelta))
                            {
                                errors.Add(FormatDestructiveTextLengthMessage(currentSlide, start.ProjectId, start.TaskId, allowedDelta, actualDelta));
                            }
                        }
                    }
                }
            }

            // 図形座標・サイズの比較
            bool exemptFullShapePosition = flags.HasFlag(PPValidationExemptFlags.ShapePosition);
            bool perSlideShapePositionExempt = UsesPerSlideShapePositionExempt(start.ProjectId, start.TaskId);
            bool onlyNewShapesExempt = IsShapePositionExemptForNewShapesOnly(start.ProjectId, start.TaskId);
            int allowedExistingChangesCount = GetAllowedExistingShapePositionChangeCount(start.ProjectId, start.TaskId);

            if (!exemptFullShapePosition || onlyNewShapesExempt || allowedExistingChangesCount >= 0 || perSlideShapePositionExempt)
            {
                int changedExistingShapesCount = 0;
                foreach (var kvp in start.ShapePositions)
                {
                    int snapshotSlide = 0;
                    var slideParts = kvp.Key.Split('_');
                    if (slideParts.Length > 0)
                        int.TryParse(slideParts[0], out snapshotSlide);

                    int currentSlide = MapSnapshotSlideToCurrent(start.ProjectId, start.TaskId, snapshotSlide, start.SlidesCount, current.SlidesCount);
                    string currentKey = slideParts.Length > 1
                        ? currentSlide + "_" + slideParts[1]
                        : kvp.Key;

                    if (!current.ShapePositions.ContainsKey(currentKey))
                        continue;

                    var cPos = current.ShapePositions[currentKey];
                    var sPos = kvp.Value;
                    if (Math.Abs(cPos.Item1 - sPos.Item1) > PositionTolerancePt ||
                        Math.Abs(cPos.Item2 - sPos.Item2) > PositionTolerancePt ||
                        Math.Abs(cPos.Item3 - sPos.Item3) > PositionTolerancePt ||
                        Math.Abs(cPos.Item4 - sPos.Item4) > PositionTolerancePt)
                    {
                        bool slideShapePositionExempt = exemptFullShapePosition
                            && (!perSlideShapePositionExempt
                                || IsShapePositionExemptForSlide(start.ProjectId, start.TaskId, currentSlide));

                        if (slideShapePositionExempt)
                        {
                            if (onlyNewShapesExempt)
                            {
                                errors.Add($"不正な図形変更: 指示外の既存図形(ID:{currentKey})の位置・サイズが変更されています。");
                            }
                            else if (allowedExistingChangesCount >= 0)
                            {
                                changedExistingShapesCount++;
                                if (changedExistingShapesCount > allowedExistingChangesCount)
                                {
                                    errors.Add($"上限超過の図形変更: 許可された数以上の既存図形(ID:{currentKey})が変更されています。");
                                }
                            }
                        }
                        else
                        {
                            errors.Add($"Shape position/size changed on Slide {currentSlide} (ID:{currentKey})");
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

            // P3-4: Slide 1 only: allow 0 or +1.
            if (projectId == 3 && taskId == 4 && slideIndex == 1)
            {
                return actualDelta == 0 || actualDelta == 1;
            }

            // P3-6: allow 0 or +3.
            if (projectId == 3 && taskId == 6)
            {
                return actualDelta == 0 || actualDelta == 3;
            }

            // P3-7: Slide 2: allow 0 or +2. Slide 1: allow 0 or +1 (section zoom side effect).
            if (projectId == 3 && taskId == 7 && slideIndex == 2)
            {
                return actualDelta == 0 || actualDelta == 2;
            }
            if (projectId == 3 && taskId == 7 && slideIndex == 1)
            {
                return actualDelta == 0 || actualDelta == 1;
            }

            // P5-5: Slide 6 only: allow 0 or -2.
            if (projectId == 5 && taskId == 5 && slideIndex == 6)
            {
                return actualDelta == 0 || actualDelta == -2;
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
            if (projectId == 3 && taskId == 4 && slideIndex == 1)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +1、実際の変化: {actualDelta}）";
            }
            if (projectId == 3 && taskId == 6)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +3、実際の変化: {actualDelta}）";
            }
            if (projectId == 3 && taskId == 7 && slideIndex == 2)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +2、実際の変化: {actualDelta}）";
            }
            if (projectId == 3 && taskId == 7 && slideIndex == 1)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または +1、実際の変化: {actualDelta}）";
            }
            if (projectId == 5 && taskId == 5 && slideIndex == 6)
            {
                return $"不正な図形操作: スライド {slideIndex} で指示外の図形の増減が検知されました（許容: 図形数の変化は 0 または -2、実際の変化: {actualDelta}）";
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

        private static bool HasTask1_8SummaryZoomExecutedGlobally()
        {
            string evidencePath = Path.Combine(Path.GetTempPath(), "mos_ppt_task_evidence.txt");
            string logPath = Path.Combine(Path.GetTempPath(), "mos_ppt_log.txt");
            return FileContainsMarker(evidencePath, "[Task1-8] SummaryZoom")
                || FileContainsMarker(logPath, "[Task1-8] SummaryZoom");
        }

        private static bool FileContainsMarker(string path, string marker)
        {
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(marker) || !File.Exists(path))
                return false;
            try
            {
                return File.ReadAllLines(path).Any(line =>
                    line != null && line.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch
            {
                return false;
            }
        }

        private static int GetProject1AdjustedSlideNumber(int logicalSlideNumber)
        {
            if (logicalSlideNumber > 1 && HasTask1_8SummaryZoomExecutedGlobally())
                return logicalSlideNumber + 1;
            return logicalSlideNumber;
        }

        private static bool IsProject1Task1_3TargetSlide(int slideIndex)
        {
            return slideIndex == GetProject1AdjustedSlideNumber(5);
        }

        private const int Project1Task1_8InsertSlideIndex = 2;
        private const int Project1Task1_1InsertAtLogical = 4;

        private static int GetProject1Task1_1OffsetAfterSlide1()
        {
            return HasTask1_8SummaryZoomExecutedGlobally() ? 1 : 0;
        }

        private static int GetProject1Task1_1InsertSlideIndex()
        {
            return Project1Task1_1InsertAtLogical + GetProject1Task1_1OffsetAfterSlide1();
        }

        private static bool IsProject1Task1_1TargetSlide(int currentSlideIndex)
        {
            return currentSlideIndex == GetProject1Task1_1InsertSlideIndex();
        }

        private static bool IsProject1Task1_1InsertApplied(int snapshotSlidesCount, int currentSlidesCount)
        {
            return currentSlidesCount == snapshotSlidesCount + 1;
        }

        private static bool IsProject1Task1_1SnapshotPostInsert(int snapshotSlidesCount, int currentSlidesCount)
        {
            return currentSlidesCount == snapshotSlidesCount;
        }

        private static bool UsesSlideIndexMapping(int projectId, int taskId)
        {
            return projectId == 1 && (taskId == 1 || taskId == 8);
        }

        private static bool IsSlidesCountValidForTask(int projectId, int taskId, int snapshotSlidesCount, int currentSlidesCount)
        {
            if (projectId == 1 && (taskId == 1 || taskId == 8))
            {
                if (taskId == 1)
                {
                    return IsProject1Task1_1InsertApplied(snapshotSlidesCount, currentSlidesCount)
                        || IsProject1Task1_1SnapshotPostInsert(snapshotSlidesCount, currentSlidesCount);
                }
                return currentSlidesCount == snapshotSlidesCount || currentSlidesCount == snapshotSlidesCount + 1;
            }
            return currentSlidesCount == snapshotSlidesCount;
        }

        private static bool IsProject1Task1_8InsertApplied(int snapshotSlidesCount, int currentSlidesCount)
        {
            return currentSlidesCount == snapshotSlidesCount + 1;
        }

        private static int MapProject1Task1_1SnapshotToCurrent(int snapshotSlideIndex)
        {
            int offset = GetProject1Task1_1OffsetAfterSlide1();
            if (snapshotSlideIndex < Project1Task1_1InsertAtLogical)
                return snapshotSlideIndex + (snapshotSlideIndex >= 2 ? offset : 0);
            return snapshotSlideIndex + 1 + offset;
        }

        private static int MapSnapshotSlideToCurrent(
            int projectId,
            int taskId,
            int snapshotSlideIndex,
            int snapshotSlidesCount,
            int currentSlidesCount)
        {
            if (projectId == 1 && taskId == 1)
            {
                if (IsProject1Task1_1InsertApplied(snapshotSlidesCount, currentSlidesCount))
                    return MapProject1Task1_1SnapshotToCurrent(snapshotSlideIndex);
                return snapshotSlideIndex;
            }
            if (projectId == 1 && taskId == 8)
            {
                if (IsProject1Task1_8InsertApplied(snapshotSlidesCount, currentSlidesCount))
                {
                    if (snapshotSlideIndex >= Project1Task1_8InsertSlideIndex)
                        return snapshotSlideIndex + 1;
                    return snapshotSlideIndex;
                }
                return snapshotSlideIndex;
            }
            return snapshotSlideIndex;
        }

        private static bool IsProject1Task1_8TargetSlide(int currentSlideIndex)
        {
            return currentSlideIndex == Project1Task1_8InsertSlideIndex;
        }

        private static bool UsesPerSlideShapePositionExempt(int projectId, int taskId)
        {
            return projectId == 1 && (taskId == 1 || taskId == 3 || taskId == 8);
        }

        private static bool IsShapePositionExemptForSlide(int projectId, int taskId, int slideIndex)
        {
            if (projectId == 1 && taskId == 1)
                return IsProject1Task1_1TargetSlide(slideIndex);
            if (projectId == 1 && taskId == 3)
                return IsProject1Task1_3TargetSlide(slideIndex);
            if (projectId == 1 && taskId == 8)
                return IsProject1Task1_8TargetSlide(slideIndex);
            return false;
        }

        private bool IsShapePositionExemptForNewShapesOnly(int projectId, int taskId)
        {
            if (projectId == 3 && (taskId == 1 || taskId == 3 || taskId == 4 || taskId == 6)) return true;
            if (projectId == 5 && (taskId == 3 || taskId == 4 || taskId == 5)) return true; // P5-3, P5-4, P5-5
            if (projectId == 6 && taskId == 3) return true; // 6-3
            if (projectId == 9 && taskId == 1) return true; // 9-1
            if (projectId == 10 && taskId == 7) return true; // 10-7
            return false;
        }

        private int GetAllowedExistingShapePositionChangeCount(int projectId, int taskId)
        {
            if (projectId == 4 && taskId == 5) return 1; // P4-5
            if (projectId == 4 && taskId == 6) return 1; // P4-6
            if (projectId == 4 && taskId == 8) return 1; // P4-8
            if (projectId == 5 && taskId == 1) return 4; // P5-1 丸4個右端揃え
            if (projectId == 5 && taskId == 2) return 1; // P5-2
            if (projectId == 3 && taskId == 5) return 1; // P3-5
            // P3-7: section zoom side effects on multiple slides — no cap (-1). ShapesCount still strict per slide.
            if (projectId == 6 && taskId == 4) return 1; // 6-4
            if (projectId == 9 && taskId == 1) return -1; // デフォルトへ (deltaで制御)
            if (projectId == 9 && taskId == 6) return 1; // 9-6
            if (projectId == 11 && taskId == 6) return 1; // 11-6
            return -1;
        }

        private int GetAllowedShapesCountDelta(int projectId, int taskId, int slideIndex)
        {
            if (projectId == 3 && taskId == 1) return slideIndex == 7 ? 0 : 0; // P3-1
            if (projectId == 3 && taskId == 3) return slideIndex == 6 ? 0 : 0; // P3-3
            if (projectId == 3 && taskId == 4) return slideIndex == 1 ? 1 : 0; // P3-4
            if (projectId == 3 && taskId == 6) return 0;                       // P3-6
            if (projectId == 3 && taskId == 7)
            {
                if (slideIndex == 2) return 2; // P3-7 section zoom x2
                if (slideIndex == 1) return 1; // P3-7 section side effect
                return 0;
            }
            if (projectId == 5 && taskId == 3) return 0;                       // P5-3
            if (projectId == 5 && taskId == 5) return slideIndex == 6 ? -2 : 0; // P5-5
            if (projectId == 6 && taskId == 3) return slideIndex == 1 ? 1 : 0; // 6-3
            if (projectId == 9 && taskId == 1) return slideIndex == 2 ? 0 : 0; // 9-1
            if (projectId == 1 && taskId == 1)
                return IsProject1Task1_1TargetSlide(slideIndex) ? int.MaxValue : 0;
            if (projectId == 1 && taskId == 3)
                return IsProject1Task1_3TargetSlide(slideIndex) ? int.MaxValue : 0;
            if (projectId == 1 && taskId == 8)
                return IsProject1Task1_8TargetSlide(slideIndex) ? int.MaxValue : 0;

            return int.MaxValue;
        }

        private int GetAllowedTextLengthDelta(int projectId, int taskId, int slideIndex)
        {
            // 9-6: URLを「お問い合わせ」に変更 (スライド1の63文字のURLが6文字の「お問い合わせ」に置き換わるため -57文字)
            if (projectId == 9 && taskId == 6) return slideIndex == 1 ? -57 : 0;
            if (projectId == 1 && taskId == 1)
                return IsProject1Task1_1TargetSlide(slideIndex) ? int.MaxValue : 0;
            if (projectId == 1 && taskId == 3)
                return IsProject1Task1_3TargetSlide(slideIndex) ? int.MaxValue : 0;
            if (projectId == 1 && taskId == 8)
                return IsProject1Task1_8TargetSlide(slideIndex) ? int.MaxValue : 0;
            if (projectId == 4 && taskId == 1)
                return slideIndex == 1 ? int.MaxValue : 0; // P4-1

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
                bool isTask11_7 = IsCurrentTask(11, 7);
                if (!isTask11_7) return;
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

        private void Glow4_3PollTimer_Tick(object sender, EventArgs e)
        {
            if (_task4_3GlowLogged) return;
            if (!IsCurrentTask(4, 3)) return;
            try
            {
                TryDetectAndLogTask4_3Glow();
            }
            catch { }
        }

        /// <summary>
        /// 4-3 から離脱する直前に光彩を即時確認し、条件一致なら証跡ログを確定する。
        /// </summary>
        private void TryLogTask4_3GlowOnTaskBoundary()
        {
            try
            {
                TryDetectAndLogTask4_3Glow();
            }
            catch { }
        }

        private bool TryDetectAndLogTask4_3Glow()
        {
            if (_task4_3GlowLogged) return true;
            if (!TryDetectTask4_3GlowOnActivePresentation())
                return false;

            Logger.LogTask4_3Glow();
            _task4_3GlowLogged = true;
            return true;
        }

        private bool TryDetectTask4_3GlowOnActivePresentation()
        {
            if (Application == null || Application.Presentations == null) return false;

            PowerPoint.Presentation pres = null;
            try
            {
                pres = Application.ActivePresentation;
                if (pres == null) return false;

                PowerPoint.Slides slides = null;
                PowerPoint.Slide slide = null;
                PowerPoint.Shapes shapes = null;
                try
                {
                    slides = pres.Slides;
                    if (slides == null || slides.Count < 1) return false;
                    slide = slides[1];
                    if (slide == null) return false;
                    shapes = slide.Shapes;
                    if (shapes != null && HasTask4_3Glow18Accent6Com(shapes))
                        return true;
                }
                finally
                {
                    if (shapes != null) try { Marshal.ReleaseComObject(shapes); } catch { }
                    if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { }
                    if (slides != null) try { Marshal.ReleaseComObject(slides); } catch { }
                }

                string tempPath = Path.Combine(Path.GetTempPath(), "mos_4_3_vsto_" + Guid.NewGuid().ToString("N") + ".pptx");
                try
                {
                    pres.SaveCopyAs(tempPath);
                    return PptxGlowOpenXmlReader.ContainsGlow18ptAccent6OnSlide(tempPath, 1);
                }
                finally
                {
                    if (File.Exists(tempPath))
                    {
                        try { File.Delete(tempPath); } catch { }
                    }
                }
            }
            finally
            {
                if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { }
            }
        }

        private static bool HasTask4_3Glow18Accent6Com(PowerPoint.Shapes shapes)
        {
            if (shapes == null) return false;
            int count = 0;
            try { count = shapes.Count; } catch { return false; }
            for (int i = 1; i <= count; i++)
            {
                PowerPoint.Shape sh = null;
                try
                {
                    sh = shapes[i];
                    if (HasTask4_3Glow18Accent6Com(sh))
                        return true;
                }
                finally
                {
                    if (sh != null) try { Marshal.ReleaseComObject(sh); } catch { }
                }
            }
            return false;
        }

        private static bool HasTask4_3Glow18Accent6Com(PowerPoint.GroupShapes group)
        {
            if (group == null) return false;
            int count = 0;
            try { count = group.Count; } catch { return false; }
            for (int i = 1; i <= count; i++)
            {
                PowerPoint.Shape sh = null;
                try
                {
                    sh = group[i];
                    if (HasTask4_3Glow18Accent6Com(sh))
                        return true;
                }
                finally
                {
                    if (sh != null) try { Marshal.ReleaseComObject(sh); } catch { }
                }
            }
            return false;
        }

        private static bool HasTask4_3Glow18Accent6Com(PowerPoint.Shape sh)
        {
            if (sh == null) return false;

            try
            {
                if (sh.Type == Office.MsoShapeType.msoGroup)
                {
                    PowerPoint.GroupShapes group = null;
                    try
                    {
                        group = sh.GroupItems;
                        return HasTask4_3Glow18Accent6Com(group);
                    }
                    finally
                    {
                        if (group != null) try { Marshal.ReleaseComObject(group); } catch { }
                    }
                }
            }
            catch { }

            if (!IsTask4_3PictureCandidate(sh))
                return false;

            dynamic glow = null;
            try
            {
                glow = sh.Glow;
                if (glow == null) return false;

                float radius = 0f;
                try { radius = (float)glow.Radius; } catch { }
                if (radius < 14f || radius > 22f) return false;

                PowerPoint.ColorFormat cf = null;
                try
                {
                    cf = glow.Color;
                    if (cf == null) return false;
                    try
                    {
                        if (cf.ObjectThemeColor == Office.MsoThemeColorIndex.msoThemeColorAccent6)
                            return true;
                    }
                    catch { }

                    try
                    {
                        int rgb = (int)cf.RGB;
                        int r = rgb & 0xFF;
                        int g = (rgb >> 8) & 0xFF;
                        int b = (rgb >> 16) & 0xFF;
                        if (r >= 60 && r <= 140 && g >= 140 && g <= 210 && b >= 40 && b <= 120)
                            return true;
                    }
                    catch { }
                }
                finally
                {
                    if (cf != null) try { Marshal.ReleaseComObject(cf); } catch { }
                }
            }
            catch { }
            finally
            {
                if (glow != null) try { Marshal.ReleaseComObject(glow); } catch { }
            }

            return false;
        }

        private static bool IsTask4_3PictureCandidate(PowerPoint.Shape sh)
        {
            if (sh == null) return false;
            try
            {
                int t = (int)sh.Type;
                if (t == (int)Office.MsoShapeType.msoPicture || t == 11)
                    return true;
                if (sh.Type != Office.MsoShapeType.msoPlaceholder)
                    return false;

                PowerPoint.PlaceholderFormat pf = null;
                try
                {
                    pf = sh.PlaceholderFormat;
                    if (pf != null && pf.ContainedType == Office.MsoShapeType.msoPicture)
                        return true;
                }
                catch { }
                finally
                {
                    if (pf != null) try { Marshal.ReleaseComObject(pf); } catch { }
                }

                string name = null;
                try { name = sh.Name; } catch { }
                name = (name ?? string.Empty).ToLowerInvariant();
                return name.Contains("picture") || name.Contains("画像") || name.Contains("図");
            }
            catch { return false; }
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
                if (!IsCurrentTask(1, 2) && !IsCurrentTask(1, 3) && !IsCurrentTask(1, 4) && !IsCurrentTask(1, 8)) return;

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

                    if (IsCurrentTask(1, 8) && !_task1_8Logged)
                    {
                        PowerPoint.Slides slides = null;
                        PowerPoint.Slide slide2 = null;
                        try
                        {
                            slides = pres.Slides;
                            if (slides != null && slides.Count >= 2)
                            {
                                slide2 = slides[2];
                                if (slide2 != null)
                                {
                                    PowerPoint.Shapes shapes = slide2.Shapes;
                                    if (shapes != null)
                                    {
                                        for (int i = 1; i <= shapes.Count; i++)
                                        {
                                            PowerPoint.Shape sh = null;
                                            try
                                            {
                                                sh = shapes[i];
                                                if (sh.HasTextFrame == Office.MsoTriState.msoTrue)
                                                {
                                                    var tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                                    string text = tf?.TextRange?.Text ?? "";
                                                    if (text.IndexOf("ご提案のポイント", StringComparison.OrdinalIgnoreCase) >= 0)
                                                    {
                                                        // ズームオブジェクト (Shape.Type == msoZoom (21)) が存在することを確認
                                                        bool hasZoom = false;
                                                        for (int j = 1; j <= shapes.Count; j++)
                                                        {
                                                            PowerPoint.Shape shZoom = null;
                                                            try
                                                            {
                                                                shZoom = shapes[j];
                                                                if ((int)shZoom.Type == 21 || shZoom.Name.Contains("Zoom") || shZoom.Name.Contains("ズーム"))
                                                                {
                                                                    hasZoom = true;
                                                                    break;
                                                                }
                                                            }
                                                            catch { }
                                                            finally { if (shZoom != null) try { Marshal.ReleaseComObject(shZoom); } catch { } }
                                                        }

                                                        if (hasZoom)
                                                        {
                                                            Logger.LogTask1_8SummaryZoom();
                                                            _task1_8Logged = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                            catch { }
                                            finally { if (sh != null) try { Marshal.ReleaseComObject(sh); } catch { } }
                                        }
                                        try { Marshal.ReleaseComObject(shapes); } catch { }
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (slide2 != null) try { Marshal.ReleaseComObject(slide2); } catch { }
                            if (slides != null) try { Marshal.ReleaseComObject(slides); } catch { }
                        }
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
            if (_glow4_3PollTimer != null)
            {
                _glow4_3PollTimer.Stop();
                _glow4_3PollTimer.Dispose();
                _glow4_3PollTimer = null;
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
