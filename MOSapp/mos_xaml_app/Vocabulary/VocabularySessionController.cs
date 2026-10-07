using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using MosPracticeClient;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace MOSExcelMogiApp.Vocabulary
{
    public sealed class VocabularySessionController : IDisposable
    {
        public enum Phase
        {
            Idle,
            Tutorial,
            Quiz,
            Finished
        }

        readonly Dispatcher _dispatcher;
        readonly Action<string> _setKeywordDisplay;
        readonly Action<string> _setProgressDisplay;
        readonly Action _onFinished;
        readonly Func<IntPtr> _getExcelHwnd;
        readonly Func<ExcelApp> _getExcelApp;
        readonly Ui.ViewModels.VocabularyUiAnchorHandler _getUiAnchor;

        VocabularyEventWatcher _watcher;
        DispatcherTimer _pollTimer;
        CoachMarkOverlayWindow _coach;
        List<VocabularyKeywordItem> _queue = new List<VocabularyKeywordItem>();
        int _index;
        bool _awaitingDismiss;
        bool _tutorialMode;
        VocabularyCategory _category;
        Phase _phase = Phase.Idle;
        readonly HashSet<string> _acceptedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>順次 detectKeys（SelectTable → TableDesignTab 等）の達成済みキー。</summary>
        readonly HashSet<string> _sequentialProgress = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        VocabularyKeywordItem _current;
        bool _currentSolved;

        /// <summary>テーブル／グラフのチュートリアル2段階目（デザインタブ）待ち。</summary>
        int _tutorialSubStep;
        DateTime _suppressSelectionPollUntil = DateTime.MinValue;
        bool _tutorialAdvanceBusy;
        /// <summary>ポーリングは「未選択→選択」の立ち上がりのみ進行（既選択でのスキップ防止）。</summary>
        bool _prevPollTargetSelected;
        /// <summary>1/2 で選択解除に失敗したとき、ポーリングでは進めずイベント待ちにする。</summary>
        bool _step0IgnorePoll;
        /// <summary>キーワード開始直後の残留選択を不正解にしない。</summary>
        DateTime _suppressWrongUntil = DateTime.MinValue;
        /// <summary>不正解案内の段階。1=範囲/グラフ選択、2=デザインタブ。</summary>
        int _wrongGuideStep;
        /// <summary>クリック校正中。</summary>
        bool _calibratingHole;
        int _calibrateCornerIndex;
        Point? _calibrateTopLeftPhysical;
        /// <summary>校正比率（物理座標は表示のたびに再計算）。</summary>
        VocabularyHighlightCalibration.HoleRatio _calibratedTableRatio;
        VocabularyHighlightCalibration.HoleRatio _calibratedChartRatio;
        /// <summary>locked 校正があるとき COM 穴に落とさない。</summary>
        bool _calibrationLocked;

        /// <summary>この問題で一度でも不正解になった（記録は1問1回）。</summary>
        bool _currentMissed;
        /// <summary>「間違えた問題だけを解く」で出題中。</summary>
        bool _mistakesOnlyMode;
        /// <summary>不正解の案内（タブとボタンを光らせる）を表示中。</summary>
        bool _wrongCoachVisible;
        string _lastWrongHolesKey;
        /// <summary>クイズ中に最後に見えた選択タブ名。</summary>
        string _lastQuizTab;
        /// <summary>選択タブを一度でも取得できた。</summary>
        bool _quizTabKnown;
        bool _onTargetTab;
        /// <summary>正しいタブのときのボタン位置（物理座標）。</summary>
        Rect? _targetControlRect;
        bool _ribbonProbeBusy;
        bool _tutorialTabProbeBusy;
        VocabularyRibbonClickWatcher _clickWatcher;
        /// <summary>ファイルタブ（バックステージ）が開いている。</summary>
        bool _backstageOpen;
        string _lastBackstagePage;
        /// <summary>不正解の案内の段階（WrongStep*）。</summary>
        int _wrongStep;

        sealed class RibbonSnapshot
        {
            public string SelectedTab;
            public Rect? TabRect;
            public Rect? ControlRect;
            public bool BackstageOpen;
            public string BackstagePage;
            public Rect? InfoRect;
            public Rect? BackRect;
        }

        public VocabularySessionController(
            Dispatcher dispatcher,
            Action<string> setKeywordDisplay,
            Action<string> setProgressDisplay,
            Action onFinished,
            Func<IntPtr> getExcelHwnd,
            Func<ExcelApp> getExcelApp = null,
            Ui.ViewModels.VocabularyUiAnchorHandler getUiAnchor = null)
        {
            _dispatcher = dispatcher;
            _setKeywordDisplay = setKeywordDisplay;
            _setProgressDisplay = setProgressDisplay;
            _onFinished = onFinished;
            _getExcelHwnd = getExcelHwnd;
            _getExcelApp = getExcelApp;
            _getUiAnchor = getUiAnchor;
        }

        public Phase CurrentPhase => _phase;
        public VocabularyKeywordItem Current => _current;
        public bool IsActive => _phase == Phase.Tutorial || _phase == Phase.Quiz;
        public bool CurrentSolved => _currentSolved;

        public void Start(VocabularyCategory category)
        {
            StopWatcher();
            VocabularyEventWatcher.ClearEvents();

            _category = category;
            _mistakesOnlyMode = false;
            var ordered = VocabularyCatalog.Filter(category).ToList();
            var quizItems = ordered.OrderBy(_ => Guid.NewGuid()).ToList();

            var tutorial = new List<VocabularyKeywordItem>();
            if (category == VocabularyCategory.TabButton || category == VocabularyCategory.Both)
            {
                var tab = ordered.FirstOrDefault(IsTableKeyword)
                          ?? quizItems.FirstOrDefault(i => !i.IsFunction);
                if (tab != null) tutorial.Add(tab);
            }
            if (category == VocabularyCategory.Function || category == VocabularyCategory.Both)
            {
                var fn = ordered.FirstOrDefault(i =>
                             string.Equals(i.FormulaName, "MAX", StringComparison.OrdinalIgnoreCase)
                             || string.Equals(i.Keyword, "最高", StringComparison.Ordinal))
                         ?? quizItems.FirstOrDefault(i => i.IsFunction);
                if (fn != null) tutorial.Add(fn);
            }

            var quiz = quizItems;
            if (tutorial.Count > 0)
            {
                var nonTable = quizItems.Where(i => !IsTableKeyword(i)).ToList();
                var tables = quizItems.Where(IsTableKeyword).ToList();
                quiz = new List<VocabularyKeywordItem>();
                if (nonTable.Count > 0)
                    quiz.Add(nonTable[0]);
                quiz.AddRange(nonTable.Skip(1).Concat(tables).OrderBy(_ => Guid.NewGuid()));
            }

            _queue = tutorial.Count > 0
                ? tutorial.Concat(quiz).ToList()
                : quizItems;

            if (_queue.Count == 0)
            {
                MessageBox.Show("出題するキーワードがありません。", "単語帳", MessageBoxButton.OK, MessageBoxImage.Information);
                _phase = Phase.Finished;
                _onFinished?.Invoke();
                return;
            }

            _index = 0;
            _tutorialMode = tutorial.Count > 0;
            _phase = _tutorialMode ? Phase.Tutorial : Phase.Quiz;
            // 前回の選択（テーブル等）をイベントにする前に切る
            WriteVocabModeFlag(false);
            VocabularyEventWatcher.ClearEvents();
            _calibrationLocked = false;
            _calibratedTableRatio = null;
            _calibratedChartRatio = null;
            TryLoadPersistedCalibration();
            EnsureA1ThenBegin();
        }

        /// <summary>単語帳を開いたら必ず A1・ホームタブにしてから出題する。</summary>
        void EnsureA1ThenBegin()
        {
            if (TryResetSheetToA1())
            {
                ArmSession();
                return;
            }

            var retry = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            int tries = 0;
            retry.Tick += (s, e) =>
            {
                if (_phase != Phase.Tutorial && _phase != Phase.Quiz)
                {
                    retry.Stop();
                    return;
                }
                tries++;
                if (TryResetSheetToA1() || tries >= 10)
                {
                    retry.Stop();
                    ArmSession();
                }
            };
            retry.Start();
        }

        void ArmSession()
        {
            if (_watcher != null) return;
            if (_phase != Phase.Tutorial && _phase != Phase.Quiz) return;
            WriteVocabModeFlag(false);
            VocabularyEventWatcher.ClearEvents();
            StartWatcher();
            ShowCurrent(showTutorialCoach: _tutorialMode);
        }

        bool TryResetSheetToA1()
        {
            try
            {
                var excel = _getExcelApp?.Invoke();
                if (excel == null) return false;
                var ws = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                if (ws == null) return false;
                ws.Range["A1"].Select();
                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
                if (hwnd == IntPtr.Zero)
                {
                    try { hwnd = new IntPtr(excel.Hwnd); } catch { }
                }
                try { VocabularyRibbonTabProbe.TryActivateHomeTab(hwnd); } catch { }
                return true;
            }
            catch
            {
                return false;
            }
        }

        void TryLoadPersistedCalibration()
        {
            try
            {
                var store = VocabularyHighlightCalibration.Load();
                if (store == null || !store.Locked) return;

                _calibrationLocked = true;
                if (store.Table != null && store.Table.IsValid)
                    _calibratedTableRatio = store.Table;
                if (store.Chart != null && store.Chart.IsValid)
                    _calibratedChartRatio = store.Chart;
            }
            catch { }
        }

        void PersistCalibration()
        {
            try
            {
                if (_calibratedTableRatio == null && _calibratedChartRatio == null)
                    return;

                var store = VocabularyHighlightCalibration.Load()
                            ?? new VocabularyHighlightCalibration.Store();
                store.Locked = true;
                if (_calibratedTableRatio != null)
                    store.Table = _calibratedTableRatio;
                if (_calibratedChartRatio != null)
                    store.Chart = _calibratedChartRatio;

                VocabularyHighlightCalibration.Save(store);
                _calibrationLocked = true;
            }
            catch { }
        }

        /// <summary>現在の Excel HWND から校正穴を解決。</summary>
        Rect? ResolveCalibratedHoleNow(bool table)
        {
            var ratio = table ? _calibratedTableRatio : _calibratedChartRatio;
            if (ratio == null || !ratio.IsValid) return null;
            IntPtr hwnd = IntPtr.Zero;
            try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
            return VocabularyHighlightCalibration.ResolveHole(ratio, hwnd);
        }

        bool HasValidCalibratedRatio(bool table)
        {
            var ratio = table ? _calibratedTableRatio : _calibratedChartRatio;
            return ratio != null && ratio.IsValid;
        }

        public void GoNext()
        {
            if (_awaitingDismiss) return;

            bool leavingTableTutorial = _phase == Phase.Tutorial
                && IsTableKeyword(_current)
                && _tutorialSubStep >= 4;

            _index++;
            if (_index >= _queue.Count)
            {
                Finish();
                return;
            }

            bool stillTutorial = _tutorialMode && _index < CountLeadingTutorial();
            _phase = stillTutorial ? Phase.Tutorial : Phase.Quiz;
            ShowCurrent(showTutorialCoach: stillTutorial);
            if (leavingTableTutorial)
                ShowLetsSolveBubble(resumeTutorial: stillTutorial);
        }

        /// <summary>テーブルチュートリアルの次へボタンのあと。OK で閉じる。</summary>
        void ShowLetsSolveBubble(bool resumeTutorial)
        {
            ShowCoach(
                title: "チュートリアル",
                message: "それでは、問題を解いてみましょう！",
                hintOverride: null,
                allowDismiss: true,
                clickThrough: false,
                onDismiss: () =>
                {
                    if (resumeTutorial && _phase == Phase.Tutorial)
                        ShowTutorialCoachStep();
                },
                holesOverride: new List<Rect>(),
                appendSelectHint: false,
                coverScreen: true,
                centerBubble: true,
                showOkButton: true);
        }

        int CountLeadingTutorial()
        {
            int n = 0;
            if (_category == VocabularyCategory.TabButton || _category == VocabularyCategory.Both) n++;
            if (_category == VocabularyCategory.Function || _category == VocabularyCategory.Both) n++;
            return Math.Min(n, _queue.Count);
        }

        void ShowCurrent(bool showTutorialCoach)
        {
            _current = _queue[_index];
            _currentSolved = false;
            _tutorialSubStep = 0;
            _tutorialAdvanceBusy = false;
            _sequentialProgress.Clear();
            _prevPollTargetSelected = true; // 既選択のまま即進行しない
            _step0IgnorePoll = false;
            _wrongGuideStep = 0;
            _currentMissed = false;
            _lastQuizTab = null;
            _quizTabKnown = false;
            _onTargetTab = false;
            _targetControlRect = null;
            _lastBackstagePage = null;
            _wrongStep = 0;
            RebuildAcceptedKeys();
            // 残留選択をイベントにする前に切り、A1・ホームへ戻してから判定を始める
            WriteVocabModeFlag(false);
            TryResetSheetToA1();
            VocabularyEventWatcher.ClearEvents();
            _suppressWrongUntil = DateTime.UtcNow.AddMilliseconds(800);
            _suppressSelectionPollUntil = _suppressWrongUntil;
            WriteVocabModeFlag(true);

            CloseCoach();
            string prefix = _phase == Phase.Tutorial ? "【チュートリアル】" : "";
            _setKeywordDisplay?.Invoke(prefix + (_current.Keyword ?? ""));
            _setProgressDisplay?.Invoke($"{_index + 1}/{_queue.Count}");

            if (showTutorialCoach && IsTableKeyword(_current))
            {
                // 1: 問題文のキーワード。セル選択の準備はクリック後。
                ShowTutorialCoachStep();
                return;
            }

            if (showTutorialCoach && IsChartKeyword(_current))
            {
                PrepareChartOrTableSelection();
                var delay = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(350) };
                delay.Tick += (s, e) =>
                {
                    delay.Stop();
                    if (_currentSolved || _tutorialSubStep != 0) return;
                    if (NeedsHoleCalibration(_current))
                        StartHoleCalibration();
                    else
                        ShowTutorialCoachStep();
                };
                delay.Start();
                return;
            }

            if (showTutorialCoach)
                ShowTutorialCoachStep();
        }

        bool NeedsHoleCalibration(VocabularyKeywordItem item)
        {
            if (item == null) return false;
            if (IsTableKeyword(item)) return !HasValidCalibratedRatio(table: true);
            if (IsChartKeyword(item)) return !HasValidCalibratedRatio(table: false);
            return false;
        }

        void StartHoleCalibration()
        {
            IntPtr hwnd = IntPtr.Zero;
            try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
            if (VocabularyHighlightCalibration.TryGetExcelWindowPhysical(hwnd) == null)
            {
                // HWND 未準備なら少し待って再試行
                var retry = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
                int tries = 0;
                retry.Tick += (s, e) =>
                {
                    tries++;
                    try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
                    if (VocabularyHighlightCalibration.TryGetExcelWindowPhysical(hwnd) != null || tries >= 8)
                    {
                        retry.Stop();
                        if (VocabularyHighlightCalibration.TryGetExcelWindowPhysical(hwnd) == null)
                        {
                            MessageBox.Show(
                                "Excel ウィンドウを取得できないため、ハイライト位置を設定できません。",
                                "単語帳",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                            ShowTutorialCoachStep();
                            return;
                        }
                        BeginHoleCalibrationUi();
                    }
                };
                retry.Start();
                return;
            }

            BeginHoleCalibrationUi();
        }

        void BeginHoleCalibrationUi()
        {
            _calibratingHole = true;
            _calibrateCornerIndex = 0;
            _calibrateTopLeftPhysical = null;
            _suppressSelectionPollUntil = DateTime.UtcNow.AddHours(1);

            bool isTable = IsTableKeyword(_current);
            ShowCoach(
                title: "ハイライト位置の設定",
                message: isTable
                    ? "テーブルの『左上』の角をクリックしてください。"
                    : "グラフの『左上』の角をクリックしてください。",
                hintOverride: null,
                allowDismiss: false,
                clickThrough: false,
                onDismiss: null,
                holesOverride: Array.Empty<Rect>(),
                appendSelectHint: false,
                beginClickCapture: true);
        }

        void OnCalibrationPhysicalClick(Point physical)
        {
            if (!_calibratingHole || _coach == null) return;

            if (_calibrateCornerIndex == 0)
            {
                _calibrateTopLeftPhysical = physical;
                _calibrateCornerIndex = 1;
                _coach.UpdateMessage(
                    "ハイライト位置の設定",
                    IsTableKeyword(_current)
                        ? "次にテーブルの『右下』の角をクリックしてください。"
                        : "次にグラフの『右下』の角をクリックしてください。");
                return;
            }

            if (!_calibrateTopLeftPhysical.HasValue) return;

            // クリック座標と同じ基準のウィンドウ矩形で比率化（オーバーレイの矩形を優先）
            Rect? win = _coach.TryGetOverlayWindowPhysical();
            if (!win.HasValue)
            {
                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
                win = VocabularyHighlightCalibration.TryGetExcelWindowPhysical(hwnd);
            }

            if (!win.HasValue)
            {
                MessageBox.Show(
                    "ウィンドウサイズを取得できないため保存できません。もう一度設定してください。",
                    "単語帳",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            var tl = _calibrateTopLeftPhysical.Value;
            double x = Math.Min(tl.X, physical.X);
            double y = Math.Min(tl.Y, physical.Y);
            double w = Math.Max(24, Math.Abs(physical.X - tl.X));
            double h = Math.Max(24, Math.Abs(physical.Y - tl.Y));
            var hole = new Rect(x, y, w, h);
            var ratio = VocabularyHighlightCalibration.ToRatio(hole, win.Value);
            if (!ratio.IsValid)
            {
                MessageBox.Show(
                    "選択範囲が小さすぎます。左上と右下をもう一度クリックしてください。",
                    "単語帳",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                _calibrateCornerIndex = 0;
                _calibrateTopLeftPhysical = null;
                _coach.ClearCalibrationMarkers();
                _coach.UpdateMessage(
                    "ハイライト位置の設定",
                    IsTableKeyword(_current)
                        ? "テーブルの『左上』の角をクリックしてください。"
                        : "グラフの『左上』の角をクリックしてください。");
                return;
            }

            if (IsTableKeyword(_current))
                _calibratedTableRatio = ratio;
            else if (IsChartKeyword(_current))
                _calibratedChartRatio = ratio;

            PersistCalibration();

            _calibratingHole = false;
            _calibrateCornerIndex = 0;
            _calibrateTopLeftPhysical = null;
            try { _coach.PhysicalClickCaptured -= OnCalibrationPhysicalClick; } catch { }
            try { _coach.EndClickCapture(); } catch { }

            _suppressSelectionPollUntil = DateTime.UtcNow.AddMilliseconds(500);
            CapturePollBaseline();
            MessageBox.Show(
                "ハイライト位置を決定しました。\n以降は画面比率で同じ位置に表示します。",
                "単語帳",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            ShowTutorialCoachStep();
        }

        int TableSelectStep => 1;
        int TableDesignStep => 2;
        int ChartSelectStep => 0;
        int ChartDesignStep => 1;

        int CurrentSelectStep => IsTableKeyword(_current) ? TableSelectStep : ChartSelectStep;
        int CurrentDesignStep => IsTableKeyword(_current) ? TableDesignStep : ChartDesignStep;

        /// <summary>セル／グラフ選択の開始。既選択のまま即進行しない。</summary>
        void PrepareChartOrTableSelection()
        {
            try
            {
                var excel = _getExcelApp?.Invoke();
                if (excel != null && IsTableKeyword(_current))
                {
                    try
                    {
                        var ws = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                        ws?.Range["A1"]?.Select();
                    }
                    catch { }
                }
            }
            catch { }

            _prevPollTargetSelected = true;
            _step0IgnorePoll = false;
            _suppressSelectionPollUntil = DateTime.UtcNow.AddMilliseconds(800);
            int selectStep = CurrentSelectStep;
            _dispatcher.BeginInvoke(new Action(() =>
            {
                if (_tutorialSubStep != selectStep || _currentSolved) return;
                CapturePollBaseline();
                if (_prevPollTargetSelected)
                    _step0IgnorePoll = true;
            }), DispatcherPriority.ApplicationIdle);
        }

        /// <summary>デザインタブ選択の開始。既にそのタブでも即正解にしない。</summary>
        void PrepareDesignTabStep()
        {
            IntPtr hwnd = IntPtr.Zero;
            try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
            bool homeActivated = false;
            try { homeActivated = VocabularyRibbonTabProbe.TryActivateHomeTab(hwnd); } catch { }

            _suppressSelectionPollUntil = DateTime.UtcNow.AddMilliseconds(300);
            if (homeActivated)
            {
                // ホームに戻せたので、以後デザインタブが選ばれていればクリックされたとみなす。
                _prevPollTargetSelected = false;
                return;
            }

            _prevPollTargetSelected = true;
            int designStep = CurrentDesignStep;
            _dispatcher.BeginInvoke(new Action(() =>
            {
                if (_tutorialSubStep != designStep || _currentSolved) return;
                CapturePollBaseline();
            }), DispatcherPriority.ApplicationIdle);
        }

        bool TryGetUiAnchor(string which, out IntPtr hwnd, out Rect rect)
        {
            hwnd = IntPtr.Zero;
            rect = Rect.Empty;
            try
            {
                if (_getUiAnchor != null && _getUiAnchor(which, out hwnd, out rect))
                    return rect.Width >= 4 && rect.Height >= 4;
            }
            catch { }
            hwnd = IntPtr.Zero;
            rect = Rect.Empty;
            return false;
        }

        void CapturePollBaseline()
        {
            try
            {
                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
                var excel = _getExcelApp?.Invoke();
                _prevPollTargetSelected = IsTutorialTargetCurrentlyMet(excel, hwnd);
            }
            catch
            {
                _prevPollTargetSelected = false;
            }
        }

        bool IsTutorialTargetCurrentlyMet(ExcelApp excel, IntPtr hwnd)
        {
            if (_tutorialSubStep == CurrentSelectStep)
            {
                if (IsTableKeyword(_current))
                    return excel != null && VocabularyHighlightHelper.IsTableCurrentlySelected(excel);
                if (IsChartKeyword(_current))
                    return excel != null && VocabularyHighlightHelper.IsChartCurrentlySelected(excel);
                return false;
            }

            if (_tutorialSubStep == CurrentDesignStep)
            {
                if (IsTableKeyword(_current))
                    return VocabularyRibbonTabProbe.IsTableDesignTabSelected(hwnd);
                if (IsChartKeyword(_current))
                    return VocabularyRibbonTabProbe.IsChartDesignTabSelected(hwnd);
            }

            return false;
        }

        void ShowTableTutorialStep()
        {
            if (_tutorialSubStep == 0)
            {
                ShowAnchoredCoach(
                    which: "keyword",
                    message: "問題文に出てくるキーワードがここに表示されます！",
                    clickThrough: false,
                    onOverlayClick: null,
                    expectedStep: 0,
                    attempt: 0,
                    showOkButton: true,
                    onOk: AdvanceFromKeywordStep);
                return;
            }

            if (_tutorialSubStep == 1)
            {
                ShowCoach(
                    title: "チュートリアル",
                    message: "ハイライトされたテーブルをクリックして選択してください。",
                    hintOverride: "Table",
                    allowDismiss: false,
                    clickThrough: true,
                    onDismiss: null);
                return;
            }

            if (_tutorialSubStep == 2)
            {
                ShowCoach(
                    title: "チュートリアル",
                    message: "リボンの『テーブルデザイン』タブをクリックしてください。",
                    hintOverride: "TableDesignTab",
                    allowDismiss: false,
                    clickThrough: true,
                    onDismiss: null);
                return;
            }

            if (_tutorialSubStep == 4)
                ShowNextButtonCoach();
        }

        void AdvanceFromKeywordStep()
        {
            if (!IsTableKeyword(_current) || _tutorialSubStep != 0 || _currentSolved) return;
            _tutorialSubStep = 1;
            CloseCoach();
            PrepareChartOrTableSelection();
            if (NeedsHoleCalibration(_current))
            {
                StartHoleCalibration();
                return;
            }
            ShowTutorialCoachStep();
        }

        void ShowTableCorrectDialog()
        {
            var dialog = new VocabularyTutorialOkWindow(CurrentDisplayText());
            dialog.OkClicked += () =>
            {
                try { dialog.Close(); } catch { }
            };
            dialog.Closed += (_, __) =>
            {
                if (!IsTableKeyword(_current) || _tutorialSubStep != 3 || _currentSolved) return;
                _tutorialSubStep = 4;
                CloseCoach();
                ShowNextButtonCoach();
            };
            dialog.Show();
            _dispatcher.BeginInvoke(new Action(() =>
            {
                if (_tutorialSubStep != 3) return;
                try { dialog.UpdateLayout(); } catch { }
                ShowCorrectDialogCoach(dialog);
                // 表示直後は窓サイズが小さいことがある。確定後にダイアログ全体をマークし直す。
                _dispatcher.BeginInvoke(new Action(() =>
                {
                    if (_tutorialSubStep != 3 || _coach == null) return;
                    try { dialog.UpdateLayout(); } catch { }
                    Rect window = dialog.TryGetWindowScreenRect();
                    if (window.Width >= 8 && window.Height >= 8)
                    {
                        try { _coach.UpdateHighlights(new List<Rect> { window }); } catch { }
                    }
                }), DispatcherPriority.Render);
            }), DispatcherPriority.Loaded);
        }

        void ShowCorrectDialogCoach(VocabularyTutorialOkWindow dialog)
        {
            Rect window = dialog.TryGetWindowScreenRect();
            ShowCoach(
                title: "チュートリアル",
                message: "正解です！こちらのOKボタンを押して次の問題に行きましょう！",
                hintOverride: null,
                allowDismiss: false,
                clickThrough: true,
                onDismiss: null,
                holesOverride: window.Width >= 8 ? new List<Rect> { window } : new List<Rect>(),
                appendSelectHint: false,
                anchorHwnd: dialog.WindowHandle,
                coverScreen: true,
                pinAboveScreen: window);
        }

        void ShowNextButtonCoach()
        {
            ShowAnchoredCoach(
                which: "next",
                message: "正解したら次の問題に行きましょう！",
                clickThrough: true,
                onOverlayClick: null,
                expectedStep: 4,
                attempt: 0);
        }

        /// <summary>アプリバー上のコントロールがレイアウトされるまで穴の取得を待つ。</summary>
        void ShowAnchoredCoach(string which, string message, bool clickThrough, Action onOverlayClick, int expectedStep, int attempt, bool showOkButton = false, Action onOk = null)
        {
            if (_tutorialSubStep != expectedStep || _currentSolved) return;
            IntPtr hwnd;
            Rect rect;
            bool ok = TryGetUiAnchor(which, out hwnd, out rect);
            if (!ok && attempt < 6)
            {
                var retry = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(120) };
                retry.Tick += (s, e) =>
                {
                    retry.Stop();
                    ShowAnchoredCoach(which, message, clickThrough, onOverlayClick, expectedStep, attempt + 1, showOkButton, onOk);
                };
                retry.Start();
                return;
            }

            IntPtr problemHwnd;
            Rect problem;
            if (!TryGetUiAnchor("keyword", out problemHwnd, out problem) || problem.Width < 8)
                problem = Rect.Empty;

            // 吹き出しはアプリバー全体ではなく、案内しているコントロールのすぐ上に置く。
            Rect pin = ok && rect.Height >= 8 ? rect : problem;

            ShowCoach(
                title: "チュートリアル",
                message: message,
                hintOverride: null,
                allowDismiss: showOkButton,
                clickThrough: clickThrough,
                onDismiss: showOkButton ? onOk : null,
                holesOverride: ok ? new List<Rect> { rect } : (problem.Width >= 8 ? new List<Rect> { problem } : new List<Rect>()),
                appendSelectHint: false,
                anchorHwnd: hwnd != IntPtr.Zero ? hwnd : problemHwnd,
                onOverlayClick: showOkButton ? null : onOverlayClick,
                coverScreen: true,
                pinAboveScreen: pin,
                showOkButton: showOkButton);
        }

        void ShowTutorialCoachStep()
        {
            if (_current == null) return;

            if (IsTableKeyword(_current))
            {
                ShowTableTutorialStep();
                return;
            }

            if (IsChartKeyword(_current))
            {
                if (_tutorialSubStep == 0)
                {
                    ShowCoach(
                        title: "チュートリアル（1/2）",
                        message: "ハイライトされたグラフをクリックして選択してください。",
                        hintOverride: "Chart",
                        allowDismiss: false,
                        clickThrough: true,
                        onDismiss: null);
                    return;
                }

                ShowCoach(
                    title: "チュートリアル（2/2）",
                    message: "リボンの『グラフのデザイン』タブをクリックしてください。",
                    hintOverride: "ChartDesignTab",
                    allowDismiss: false,
                    clickThrough: true,
                    onDismiss: null);
                return;
            }

            ShowCoach(
                title: "チュートリアル",
                message: (_current.CoachMessage ?? ("キーワード「" + _current.Keyword + "」の場所を探しましょう。"))
                         + "\n（正しい場所を操作すると次へ進みます）",
                hintOverride: null,
                allowDismiss: false,
                clickThrough: true,
                onDismiss: null);
        }

        static bool NeedsTwoStepTutorial(VocabularyKeywordItem item)
        {
            if (item == null) return false;
            return IsTableKeyword(item) || IsChartKeyword(item);
        }

        static bool IsTableKeyword(VocabularyKeywordItem item)
        {
            return item != null && (
                string.Equals(item.Id, "tab_table", StringComparison.OrdinalIgnoreCase)
                || (item.Keyword ?? "").Contains("テーブル")
                || string.Equals(item.HighlightHint, "Table", StringComparison.OrdinalIgnoreCase)
                || string.Equals(item.HighlightHint, "TableThenDesignTab", StringComparison.OrdinalIgnoreCase));
        }

        static bool IsChartKeyword(VocabularyKeywordItem item)
        {
            return item != null && (
                string.Equals(item.Id, "tab_chart", StringComparison.OrdinalIgnoreCase)
                || (item.Keyword ?? "").Contains("グラフ")
                || string.Equals(item.HighlightHint, "Chart", StringComparison.OrdinalIgnoreCase)
                || string.Equals(item.HighlightHint, "ChartThenDesignTab", StringComparison.OrdinalIgnoreCase));
        }

        bool UsesSequentialDetect()
        {
            return _current?.DetectKeys != null
                   && _current.DetectKeys.Count >= 2
                   && (IsTableKeyword(_current) || IsChartKeyword(_current));
        }

        void RebuildAcceptedKeys()
        {
            _acceptedKeys.Clear();
            if (_current?.DetectKeys == null) return;
            foreach (var k in _current.DetectKeys)
            {
                if (!string.IsNullOrWhiteSpace(k))
                    _acceptedKeys.Add(k.Trim());
            }
        }

        void StartWatcher()
        {
            _watcher = new VocabularyEventWatcher();
            _watcher.EventReceived += OnVocabEvent;
            _pollTimer = new DispatcherTimer(DispatcherPriority.Background, _dispatcher)
            {
                Interval = TimeSpan.FromMilliseconds(300)
            };
            _pollTimer.Tick += (_, __) =>
            {
                _watcher?.Drain();
                PollExcelState();
            };
            _pollTimer.Start();

            _clickWatcher = new VocabularyRibbonClickWatcher(_dispatcher);
            _clickWatcher.LeftClick += OnGlobalLeftClick;
            _clickWatcher.Start();
        }

        /// <summary>
        /// テーブル選択・デザインタブ選択を Excel / UIA で直接見て進行する。
        /// </summary>
        void PollExcelState()
        {
            if (!IsActive || _current == null || _currentSolved) return;
            if (_calibratingHole) return;
            if (DateTime.UtcNow < _suppressSelectionPollUntil) return;

            try
            {
                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
                var excel = _getExcelApp?.Invoke();

                if (_phase == Phase.Tutorial && NeedsTwoStepTutorial(_current))
                {
                    if (IsTableKeyword(_current) && (_tutorialSubStep == 0 || _tutorialSubStep >= 3))
                        return;

                    if (_tutorialSubStep == CurrentDesignStep)
                    {
                        BeginTutorialDesignProbe(hwnd);
                        return;
                    }

                    bool now = IsTutorialTargetCurrentlyMet(excel, hwnd);

                    if (_tutorialSubStep == CurrentSelectStep && _step0IgnorePoll)
                    {
                        // 選択が一度外れたら立ち上がり検知に切り替え
                        if (!now)
                        {
                            _step0IgnorePoll = false;
                            _prevPollTargetSelected = false;
                        }
                        return;
                    }

                    bool rising = now && !_prevPollTargetSelected;
                    _prevPollTargetSelected = now;
                    if (!rising) return;

                    if (_tutorialSubStep == CurrentSelectStep)
                    {
                        if (IsTableKeyword(_current))
                            TryAdvanceTutorialBySelection("SelectTable");
                        else if (IsChartKeyword(_current))
                            TryAdvanceTutorialBySelection("SelectChart");
                    }
                    else if (_tutorialSubStep == CurrentDesignStep)
                    {
                        if (IsTableKeyword(_current))
                            TryAdvanceTutorialBySelection("TableDesignTab");
                        else if (IsChartKeyword(_current))
                            TryAdvanceTutorialBySelection("ChartDesignTab");
                    }
                    return;
                }

                // クイズ: 順次条件をポーリングでも進める
                if (_phase == Phase.Quiz && UsesSequentialDetect())
                {
                    if (excel != null && IsTableKeyword(_current)
                        && VocabularyHighlightHelper.IsTableCurrentlySelected(excel))
                    {
                        ApplyQuizKey("SelectTable");
                    }
                    if (excel != null && IsChartKeyword(_current)
                        && VocabularyHighlightHelper.IsChartCurrentlySelected(excel))
                    {
                        ApplyQuizKey("SelectChart");
                    }

                    if (IsTableKeyword(_current) && VocabularyRibbonTabProbe.IsTableDesignTabSelected(hwnd))
                        ApplyQuizKey("TableDesignTab");
                    if (IsChartKeyword(_current) && VocabularyRibbonTabProbe.IsChartDesignTabSelected(hwnd))
                        ApplyQuizKey("ChartDesignTab");
                }

                if (_phase == Phase.Quiz && UsesRibbonJudge() && !_currentSolved)
                    BeginRibbonProbe(hwnd);
            }
            catch { }
        }

        /// <summary>チュートリアルのデザインタブ判定。UIA は重いのでバックグラウンドで読む。</summary>
        void BeginTutorialDesignProbe(IntPtr hwnd)
        {
            if (_tutorialTabProbeBusy || hwnd == IntPtr.Zero) return;
            _tutorialTabProbeBusy = true;
            var item = _current;
            int step = _tutorialSubStep;
            bool table = IsTableKeyword(item);

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    return table
                        ? VocabularyRibbonTabProbe.IsTableDesignTabSelected(hwnd)
                        : VocabularyRibbonTabProbe.IsChartDesignTabSelected(hwnd);
                }
                catch { return false; }
            }).ContinueWith(t =>
            {
                _dispatcher.BeginInvoke(new Action(() =>
                {
                    _tutorialTabProbeBusy = false;
                    if (t.Status != System.Threading.Tasks.TaskStatus.RanToCompletion) return;
                    if (_phase != Phase.Tutorial || _current != item || _tutorialSubStep != step || _currentSolved) return;
                    if (DateTime.UtcNow < _suppressSelectionPollUntil) return;

                    bool now = t.Result;
                    bool rising = now && !_prevPollTargetSelected;
                    _prevPollTargetSelected = now;
                    if (!rising) return;
                    TryAdvanceTutorialBySelection(table ? "TableDesignTab" : "ChartDesignTab");
                }));
            });
        }

        /// <summary>タブとボタンで答える問題か（関数以外）。</summary>
        bool UsesRibbonJudge()
        {
            return _current != null
                   && !_current.IsFunction
                   && !string.IsNullOrWhiteSpace(_current.TargetTab);
        }

        bool IsFileTarget()
        {
            return VocabularyRibbonTabProbe.TabNameEquals(_current?.TargetTab, "ファイル");
        }

        bool IsTabOnlyQuestion()
        {
            return !UsesSequentialDetect() && string.IsNullOrWhiteSpace(_current?.TargetControl);
        }

        /// <summary>UIA はバックグラウンドで読み、結果だけ UI スレッドで判定する。</summary>
        void BeginRibbonProbe(IntPtr hwnd)
        {
            if (_ribbonProbeBusy || hwnd == IntPtr.Zero) return;
            _ribbonProbeBusy = true;
            var item = _current;
            string targetTab = item.TargetTab;
            string targetControl = item.TargetControl;
            bool fileTarget = IsFileTarget();
            bool wantRects = _wrongCoachVisible;

            System.Threading.Tasks.Task.Run(() =>
            {
                var snap = new RibbonSnapshot();
                try
                {
                    snap.BackstageOpen = VocabularyRibbonTabProbe.IsBackstageOpen(hwnd);
                    if (snap.BackstageOpen)
                    {
                        snap.BackstagePage = VocabularyRibbonTabProbe.TryGetBackstageSelectedPage(hwnd);
                        if (fileTarget)
                            snap.InfoRect = VocabularyRibbonTabProbe.TryGetBackstageItemScreenRect(
                                hwnd, string.IsNullOrWhiteSpace(targetControl) ? "情報" : targetControl);
                        else if (wantRects)
                            snap.BackRect = VocabularyRibbonTabProbe.TryGetBackstageItemScreenRect(hwnd, "戻る");
                        return snap;
                    }

                    snap.SelectedTab = VocabularyRibbonTabProbe.TryGetSelectedTabName(hwnd);
                    if (wantRects)
                        snap.TabRect = VocabularyRibbonTabProbe.TryGetTabScreenRectLoose(hwnd, targetTab);
                    bool onTarget = !fileTarget
                        && VocabularyRibbonTabProbe.TabNameEquals(snap.SelectedTab, targetTab);
                    if (onTarget && !string.IsNullOrWhiteSpace(targetControl))
                        snap.ControlRect = VocabularyRibbonTabProbe.TryGetRibbonControlScreenRect(hwnd, targetControl);
                }
                catch { }
                return snap;
            }).ContinueWith(t =>
            {
                _dispatcher.BeginInvoke(new Action(() =>
                {
                    _ribbonProbeBusy = false;
                    if (t.Status != System.Threading.Tasks.TaskStatus.RanToCompletion) return;
                    if (!IsActive || _phase != Phase.Quiz || _current != item || _currentSolved) return;
                    if (DateTime.UtcNow < _suppressSelectionPollUntil) return;
                    OnRibbonSnapshot(t.Result);
                }));
            });
        }

        void OnRibbonSnapshot(RibbonSnapshot snap)
        {
            bool fileTarget = IsFileTarget();
            string target = UsesSequentialDetect()
                ? (IsTableKeyword(_current) ? "テーブルデザイン" : "グラフのデザイン")
                : _current.TargetTab;

            bool backstageOpened = snap.BackstageOpen && !_backstageOpen;
            if (snap.BackstageOpen != _backstageOpen)
                LogQuiz("backstage " + (snap.BackstageOpen ? "open" : "closed") + " page=" + (snap.BackstagePage ?? "-"));
            _backstageOpen = snap.BackstageOpen;

            if (snap.BackstageOpen)
            {
                if (fileTarget)
                {
                    _targetControlRect = snap.InfoRect;
                    string page = snap.BackstagePage;
                    bool infoSelected = IsInfoPage(page);
                    if (infoSelected && !_awaitingDismiss)
                    {
                        HandleButtonCorrect();
                        return;
                    }
                    bool pageChanged = !backstageOpened
                        && !string.IsNullOrEmpty(page)
                        && !string.Equals(page, _lastBackstagePage, StringComparison.Ordinal);
                    _lastBackstagePage = page;
                    if (pageChanged && !_awaitingDismiss)
                    {
                        ShowWrongCoach();
                        return;
                    }
                }
                else
                {
                    _targetControlRect = null;
                    if (backstageOpened && !_awaitingDismiss)
                    {
                        ShowWrongCoach();
                        return;
                    }
                }
                UpdateWrongCoach(snap);
                return;
            }
            _lastBackstagePage = null;

            if (!string.IsNullOrEmpty(snap.SelectedTab))
            {
                string sel = snap.SelectedTab;
                bool changed = _lastQuizTab == null
                    ? !VocabularyRibbonTabProbe.IsHomeTabName(sel)
                    : !string.Equals(sel, _lastQuizTab, StringComparison.Ordinal);
                _lastQuizTab = sel;
                _quizTabKnown = true;
                _onTargetTab = VocabularyRibbonTabProbe.TabNameEquals(sel, target);

                if (changed && !_awaitingDismiss)
                {
                    if (_onTargetTab)
                    {
                        if (IsTabOnlyQuestion())
                        {
                            MarkCorrect();
                            return;
                        }
                    }
                    else if (!VocabularyRibbonTabProbe.IsHomeTabName(sel))
                    {
                        ShowWrongCoach();
                        return;
                    }
                }
            }

            _targetControlRect = (!fileTarget && _onTargetTab) ? snap.ControlRect : null;
            UpdateWrongCoach(snap);
        }

        const int WrongStepTab = 1;
        const int WrongStepButton = 2;
        const int WrongStepFileTab = 3;
        const int WrongStepBackstageInfo = 4;
        const int WrongStepBackstageBack = 5;

        static bool IsInfoPage(string page)
        {
            return string.Equals(page, "情報", StringComparison.Ordinal)
                   || string.Equals(page, "Info", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>不正解の案内で、いま光らせる段階。</summary>
        int CurrentWrongStep()
        {
            if (_backstageOpen)
                return IsFileTarget() ? WrongStepBackstageInfo : WrongStepBackstageBack;
            if (IsFileTarget()) return WrongStepFileTab;
            if (_onTargetTab && !string.IsNullOrWhiteSpace(_current?.TargetControl)) return WrongStepButton;
            return WrongStepTab;
        }

        /// <summary>CSV の正解（例: 挿入タブ/リンクボタン）の左右。無ければ JSON の値から作る。</summary>
        string AnswerPart(int index)
        {
            var parts = (_current?.Answer ?? "").Split('/');
            if (parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[index]))
                return parts[index].Trim();
            if (index == 0)
            {
                string tab = (_current?.TargetTab ?? "").Trim();
                return tab.EndsWith("タブ", StringComparison.Ordinal) ? tab : tab + "タブ";
            }
            return (_current?.TargetControl ?? "").Trim();
        }

        string WrongStepMessage(int step)
        {
            string tab = (_current?.TargetTab ?? "").Trim();
            switch (step)
            {
                case WrongStepButton:
                    return "『" + AnswerPart(1) + "』をクリックしてください。";
                case WrongStepFileTab:
                    return "『ファイル』タブをクリックしてください。";
                case WrongStepBackstageInfo:
                    return "『" + (string.IsNullOrWhiteSpace(_current?.TargetControl) ? "情報" : _current.TargetControl.Trim()) + "』をクリックしてください。";
                case WrongStepBackstageBack:
                    return "左上の『←』で戻り、『" + tab + "』タブをクリックしてください。";
                default:
                    return "『" + AnswerPart(0) + "』をクリックしてください。";
            }
        }

        /// <summary>段階に合う位置（UI スレッドで UIA を読む。最初の表示用）。</summary>
        List<Rect> BuildWrongStepHoles(int step)
        {
            var holes = new List<Rect>();
            IntPtr hwnd = IntPtr.Zero;
            try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
            Rect? rect = null;
            switch (step)
            {
                case WrongStepButton:
                    rect = VocabularyRibbonTabProbe.TryGetRibbonControlScreenRect(hwnd, _current.TargetControl);
                    break;
                case WrongStepFileTab:
                case WrongStepTab:
                    rect = VocabularyRibbonTabProbe.TryGetTabScreenRectLoose(hwnd, _current.TargetTab);
                    break;
                case WrongStepBackstageInfo:
                    rect = VocabularyRibbonTabProbe.TryGetBackstageItemScreenRect(
                        hwnd, string.IsNullOrWhiteSpace(_current.TargetControl) ? "情報" : _current.TargetControl);
                    break;
                case WrongStepBackstageBack:
                    rect = VocabularyRibbonTabProbe.TryGetBackstageItemScreenRect(hwnd, "戻る");
                    break;
            }
            if (rect.HasValue) holes.Add(rect.Value);
            return holes;
        }

        static string HolesKey(List<Rect> holes)
        {
            return string.Join("|", holes.Select(h => $"{h.X:0},{h.Y:0},{h.Width:0},{h.Height:0}"));
        }

        /// <summary>見回りの結果で、不正解の吹き出しの文面と光らせる位置を更新する。</summary>
        void UpdateWrongCoach(RibbonSnapshot snap)
        {
            if (!_wrongCoachVisible || _coach == null || UsesSequentialDetect()) return;

            int step = CurrentWrongStep();
            Rect? rect = null;
            switch (step)
            {
                case WrongStepButton: rect = snap.ControlRect; break;
                case WrongStepTab:
                case WrongStepFileTab: rect = snap.TabRect; break;
                case WrongStepBackstageInfo: rect = snap.InfoRect; break;
                case WrongStepBackstageBack: rect = snap.BackRect; break;
            }

            if (step != _wrongStep)
            {
                _wrongStep = step;
                _lastWrongHolesKey = null;
                try { _coach.UpdateMessage("不正解です。", WrongStepMessage(step)); } catch { }
                LogQuiz($"wrong step={step} rect={(rect.HasValue ? HolesKey(new List<Rect> { rect.Value }) : "none")}");
            }

            if (!rect.HasValue) return;
            var holes = new List<Rect> { rect.Value };
            string key = HolesKey(holes);
            if (key == _lastWrongHolesKey) return;
            if (_lastWrongHolesKey == null)
                LogQuiz($"wrong step={step} found {key}");
            _lastWrongHolesKey = key;
            try { _coach.UpdateHighlights(holes); } catch { }
        }

        static void LogQuiz(string line)
        {
            try
            {
                File.AppendAllText(
                    Path.Combine(Path.GetTempPath(), "mos_vocab_quiz_debug.txt"),
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + line + Environment.NewLine);
            }
            catch { }
        }

        /// <summary>ボタン位置のクリックを、ボタンを押したとみなす。</summary>
        void OnGlobalLeftClick(Point physical)
        {
            if (!IsActive || _phase != Phase.Quiz || _current == null || _currentSolved) return;
            if (_awaitingDismiss || !UsesRibbonJudge() || UsesSequentialDetect()) return;
            if (DateTime.UtcNow < _suppressWrongUntil) return;
            var rect = _targetControlRect;
            if (!rect.HasValue || !rect.Value.Contains(physical)) return;
            HandleButtonCorrect();
        }

        /// <summary>ボタン操作で正解。ダイアログが開いたら閉じる案内まで出す。</summary>
        void HandleButtonCorrect()
        {
            if (_currentSolved) return;
            _currentSolved = true;
            RemoveMistakeIfSolvedCleanly();
            CloseCoach();

            var item = _current;
            IntPtr hwnd = IntPtr.Zero;
            try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
            int tries = 0;
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
            timer.Tick += (s, e) =>
            {
                tries++;
                if (!IsActive || _current != item)
                {
                    timer.Stop();
                    return;
                }
                IntPtr dialog = VocabularyExcelDialogProbe.FindDialog(hwnd);
                if (dialog != IntPtr.Zero)
                {
                    timer.Stop();
                    ShowCorrectBubble(() => ShowButtonOkBubble(dialog));
                    return;
                }
                if (tries >= 10)
                {
                    timer.Stop();
                    if (_backstageOpen)
                        ShowCorrectBubble(() => CloseBackstage(hwnd));
                    else
                        ShowCorrectBubble();
                }
            };
            timer.Start();
        }

        /// <summary>次の問題の前に、開いたままのファイルタブを Esc で閉じる。</summary>
        void CloseBackstage(IntPtr hwnd)
        {
            try
            {
                if (VocabularyRibbonTabProbe.IsBackstageOpen(hwnd))
                    VocabularyExcelDialogProbe.SendEscape(hwnd);
            }
            catch { }
            _backstageOpen = false;
        }

        void ShowButtonOkBubble(IntPtr dialog)
        {
            if (!IsActive) return;
            ShowCoach(
                title: "正解！",
                message: "ボタンは合っています。次の問題に行きましょう！",
                hintOverride: null,
                allowDismiss: true,
                clickThrough: false,
                onDismiss: () => VocabularyExcelDialogProbe.CloseWithEscape(dialog),
                holesOverride: new List<Rect>(),
                appendSelectHint: false,
                coverScreen: true,
                centerBubble: true,
                showOkButton: true);
        }

        void RecordMiss()
        {
            if (_phase != Phase.Quiz || _current == null || _currentMissed) return;
            _currentMissed = true;
            try { VocabularyMistakeStore.Record(_current); } catch { }
        }

        void RemoveMistakeIfSolvedCleanly()
        {
            if (_phase != Phase.Quiz || !_mistakesOnlyMode || _currentMissed || _current == null) return;
            try { VocabularyMistakeStore.Remove(_current); } catch { }
        }

        void StopWatcher()
        {
            try { _pollTimer?.Stop(); } catch { }
            _pollTimer = null;
            if (_clickWatcher != null)
            {
                _clickWatcher.LeftClick -= OnGlobalLeftClick;
                _clickWatcher.Dispose();
                _clickWatcher = null;
            }
            if (_watcher != null)
            {
                _watcher.EventReceived -= OnVocabEvent;
                _watcher.Dispose();
                _watcher = null;
            }
        }

        void OnVocabEvent(string key)
        {
            if (!IsActive || _current == null) return;
            if (string.IsNullOrWhiteSpace(key)) return;

            _dispatcher.BeginInvoke(new Action(() =>
            {
                if (!IsActive || _current == null) return;
                if (DateTime.UtcNow < _suppressWrongUntil) return;

                // チュートリアルは案内どおりに進める。途中の操作を不正解にはしない。
                if (_phase == Phase.Tutorial)
                {
                    TryAdvanceTutorialBySelection(key);
                    return;
                }

                if (_awaitingDismiss && _phase != Phase.Tutorial)
                    return;

                if (_phase == Phase.Quiz)
                {
                    ApplyQuizKey(key);
                    return;
                }

                // 非チュートリアルで単一キーの場合
                if (IsMatch(key))
                    MarkCorrect();
                else if (IsRelevantWrongAttempt(key))
                    ShowWrongCoach();
            }));
        }

        void ApplyQuizKey(string key)
        {
            if (_currentSolved || string.IsNullOrWhiteSpace(key)) return;

            if (UsesSequentialDetect())
            {
                var keys = _current.DetectKeys;
                string first = keys[0];
                string second = keys[1];

                if (string.Equals(key, first, StringComparison.OrdinalIgnoreCase))
                {
                    if (_sequentialProgress.Add(first))
                    {
                        _wrongGuideStep = 2;
                        ShowDesignTabCoach(asWrong: false);
                    }
                    return;
                }

                if (string.Equals(key, second, StringComparison.OrdinalIgnoreCase))
                {
                    if (_sequentialProgress.Contains(first))
                    {
                        MarkCorrect();
                        return;
                    }

                    // デザインタブだけ先に来た場合はテーブル選択を促す
                    ShowWrongCoach();
                    return;
                }

                if (IsRelevantWrongAttempt(key))
                    ShowWrongCoach();
                return;
            }

            if (IsMatch(key))
            {
                if (_current.IsFunction)
                {
                    MarkCorrect();
                    return;
                }
                // 正しいタブを選ぶ前のボタン操作（ホームの同じボタン等）は不正解。
                if (!IsFileTarget() && !IsSelectedTabTargetNow())
                {
                    ShowWrongCoach();
                    return;
                }
                HandleButtonCorrect();
            }
            else if (IsRelevantWrongAttempt(key))
                ShowWrongCoach();
        }

        /// <summary>今のタブが正解タブか。ポーリングが古いときは UIA で読み直す。取れなければ true。</summary>
        bool IsSelectedTabTargetNow()
        {
            if (_onTargetTab) return true;
            IntPtr hwnd = IntPtr.Zero;
            try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
            string sel = VocabularyRibbonTabProbe.TryGetSelectedTabName(hwnd);
            if (string.IsNullOrEmpty(sel))
                return !_quizTabKnown;
            return VocabularyRibbonTabProbe.TabNameEquals(sel, _current.TargetTab);
        }

        void ShowWrongCoach()
        {
            RecordMiss();
            if (IsTableKeyword(_current) || IsChartKeyword(_current))
            {
                bool onDesign = _wrongGuideStep >= 2
                    || (_phase == Phase.Tutorial && _tutorialSubStep == CurrentDesignStep);
                _wrongGuideStep = onDesign ? 2 : 1;
                if (_wrongGuideStep >= 2)
                    ShowDesignTabCoach(asWrong: true);
                else
                    ShowSelectTargetCoach();
                return;
            }

            string hint = _current.HighlightHint;
            if (_current.IsFunction
                || string.Equals(hint, "FormulaBar", StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrWhiteSpace(_current.TargetTab))
            {
                string msg = _current.CoachMessage ?? ("正解は「" + _current.Answer + "」です。");
                ShowCoach(
                    title: "不正解です。",
                    message: msg + "\n（正しい場所を選択／操作すると進みます）",
                    hintOverride: _current.IsFunction ? "FormulaBar" : hint,
                    allowDismiss: false,
                    clickThrough: true,
                    onDismiss: null);
                return;
            }

            // 正しいタブ → ボタンの順に、いまの段階の場所だけを光らせる。位置は見回りで取り直す。
            int step = CurrentWrongStep();
            var holes = BuildWrongStepHoles(step);
            ShowCoach(
                title: "不正解です。",
                message: WrongStepMessage(step),
                hintOverride: null,
                allowDismiss: false,
                clickThrough: true,
                onDismiss: null,
                holesOverride: holes,
                appendSelectHint: false);
            _wrongCoachVisible = _coach != null;
            _wrongStep = step;
            _lastWrongHolesKey = holes.Count > 0 ? HolesKey(holes) : null;
            LogQuiz($"wrong show keyword={_current.Keyword} step={step} tab={_lastQuizTab ?? "-"} backstage={_backstageOpen} holes={(holes.Count > 0 ? HolesKey(holes) : "none")}");
        }

        void ShowSelectTargetCoach()
        {
            bool table = IsTableKeyword(_current);
            ShowCoach(
                title: "不正解です。",
                message: table
                    ? "テーブルの範囲を選択してください。"
                    : "グラフを選択してください。",
                hintOverride: table ? "Table" : "Chart",
                allowDismiss: false,
                clickThrough: true,
                onDismiss: null);
        }

        void ShowDesignTabCoach(bool asWrong)
        {
            bool table = IsTableKeyword(_current);
            ShowCoach(
                title: asWrong ? "不正解です。" : "次の操作",
                message: table
                    ? "リボンの『テーブルデザイン』タブをクリックしてください。"
                    : "リボンの『グラフのデザイン』タブをクリックしてください。",
                hintOverride: table ? "TableDesignTab" : "ChartDesignTab",
                allowDismiss: false,
                clickThrough: true,
                onDismiss: null);
        }

        void MarkCorrect()
        {
            if (_currentSolved) return;
            _currentSolved = true;
            RemoveMistakeIfSolvedCleanly();
            CloseCoach();
            ShowCorrectBubble();
        }

        /// <summary>正解はコーチマークのテキストボックスに、CSVの表示テキストを出す。</summary>
        void ShowCorrectBubble(Action onDismiss = null)
        {
            ShowCoach(
                title: "正解！",
                message: CurrentDisplayText(),
                hintOverride: null,
                allowDismiss: true,
                clickThrough: false,
                onDismiss: onDismiss,
                holesOverride: new List<Rect>(),
                appendSelectHint: false,
                coverScreen: true,
                centerBubble: true,
                showOkButton: true);
        }

        string CurrentDisplayText()
        {
            if (!string.IsNullOrWhiteSpace(_current?.DisplayText))
                return _current.DisplayText;
            return _current?.Keyword ?? "正解！";
        }

        /// <summary>チュートリアル中の選択でステップ進行。処理したら true。</summary>
        bool TryAdvanceTutorialBySelection(string key)
        {
            if (_currentSolved || _tutorialAdvanceBusy) return true;
            if (_calibratingHole) return false;
            if (NeedsTwoStepTutorial(_current) && DateTime.UtcNow < _suppressSelectionPollUntil)
                return false;

            if (!NeedsTwoStepTutorial(_current))
            {
                if (IsMatch(key))
                {
                    MarkCorrect();
                    return true;
                }
                return false;
            }

            // テーブル／グラフの選択
            if (_tutorialSubStep == CurrentSelectStep)
            {
                bool selectOk =
                    (IsTableKeyword(_current) && string.Equals(key, "SelectTable", StringComparison.OrdinalIgnoreCase))
                    || (IsChartKeyword(_current) && string.Equals(key, "SelectChart", StringComparison.OrdinalIgnoreCase));

                if (!selectOk) return false;

                _tutorialAdvanceBusy = true;
                _tutorialSubStep = CurrentDesignStep;
                _step0IgnorePoll = false;
                CloseCoach();
                PrepareDesignTabStep();
                ShowTutorialCoachStep();
                _tutorialAdvanceBusy = false;
                return true;
            }

            // デザインタブ選択
            if (_tutorialSubStep == CurrentDesignStep)
            {
                bool tabOk =
                    (IsTableKeyword(_current) && string.Equals(key, "TableDesignTab", StringComparison.OrdinalIgnoreCase))
                    || (IsChartKeyword(_current) && string.Equals(key, "ChartDesignTab", StringComparison.OrdinalIgnoreCase));

                if (!tabOk) return false;

                _tutorialAdvanceBusy = false;
                CloseCoach();
                if (IsTableKeyword(_current))
                {
                    _tutorialSubStep = 3;
                    ShowTableCorrectDialog();
                    return true;
                }

                MarkCorrect();
                return true;
            }

            return false;
        }

        bool IsMatch(string key)
        {
            if (UsesSequentialDetect())
            {
                // 順次は ApplyQuizKey 側で完了判定
                return false;
            }

            if (_acceptedKeys.Contains(key)) return true;

            if (key.StartsWith("Formula:", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrEmpty(_current.FormulaName))
            {
                string name = key.Substring("Formula:".Length);
                return string.Equals(name, _current.FormulaName, StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        bool IsRelevantWrongAttempt(string key)
        {
            if (key.StartsWith("Formula:", StringComparison.OrdinalIgnoreCase))
                return _current.IsFunction;
            // セルや図の選択は、テーブル／グラフの問題以外では答えとみなさない。
            if (key.StartsWith("Select", StringComparison.OrdinalIgnoreCase))
                return UsesSequentialDetect();
            if (key.IndexOf("Tab", StringComparison.OrdinalIgnoreCase) >= 0)
                return !_current.IsFunction;
            return !_current.IsFunction;
        }

        void ShowCoach(
            string title,
            string message,
            string hintOverride,
            bool allowDismiss,
            bool clickThrough,
            Action onDismiss,
            IReadOnlyList<Rect> holesOverride = null,
            bool appendSelectHint = true,
            bool beginClickCapture = false,
            IntPtr anchorHwnd = default,
            Action onOverlayClick = null,
            bool coverScreen = false,
            Rect bubbleScreen = default,
            Rect bubbleAboveScreen = default,
            bool centerBubble = false,
            Rect pinAboveScreen = default,
            bool showOkButton = false)
        {
            try
            {
                CloseCoach();
                _awaitingDismiss = allowDismiss;
                IntPtr excelHwnd = IntPtr.Zero;
                try { excelHwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
                IntPtr hwnd = anchorHwnd != IntPtr.Zero ? anchorHwnd : excelHwnd;

                string hint = hintOverride ?? _current?.HighlightHint;
                ExcelApp excel = null;
                try { excel = _getExcelApp?.Invoke(); } catch { }

                var holes = BuildCoachHoles(excelHwnd, excel, hint, holesOverride);

                _coach = new CoachMarkOverlayWindow();
                Action dismissAction = () =>
                {
                    _awaitingDismiss = false;
                    try { _coach.PhysicalClickCaptured -= OnCalibrationPhysicalClick; } catch { }
                    _coach = null;
                    try
                    {
                        var xl = _getExcelApp?.Invoke();
                        VocabularyHighlightHelper.ClearNativeTableHighlight(xl);
                    }
                    catch { }
                    onDismiss?.Invoke();
                };
                _coach.Dismissed += dismissAction;
                _coach.ShowCoachMark(hwnd, holes, title, message, allowDismiss, clickThrough, appendSelectHint, coverScreen, bubbleScreen, bubbleAboveScreen, centerBubble, pinAboveScreen, showOkButton);

                if (beginClickCapture)
                {
                    _coach.PhysicalClickCaptured += OnCalibrationPhysicalClick;
                    _coach.BeginClickCapture();
                }
                else if (onOverlayClick != null)
                {
                    _coach.BeginSingleClick(onOverlayClick);
                }
                else if (ShouldRetryCalibratedHole(hint, holes))
                {
                    ScheduleCalibratedHoleRetry(hint);
                }
                else if (holes.Count == 0 && holesOverride == null && !string.IsNullOrEmpty(hint))
                {
                    ScheduleLiveHoleRetry(hint);
                }
            }
            catch (Exception ex)
            {
                _awaitingDismiss = false;
                _calibratingHole = false;
                System.Diagnostics.Debug.WriteLine("[ShowCoach] " + ex.Message);
                MessageBox.Show((title ?? "") + "\n\n" + (message ?? ""), "単語帳", MessageBoxButton.OK, MessageBoxImage.Information);
                onDismiss?.Invoke();
            }
        }

        List<Rect> BuildCoachHoles(IntPtr hwnd, ExcelApp excel, string hint, IReadOnlyList<Rect> holesOverride)
        {
            var holes = new List<Rect>();
            if (holesOverride != null)
            {
                holes.AddRange(holesOverride);
                try { VocabularyHighlightHelper.ClearNativeTableHighlight(excel); } catch { }
                return holes;
            }

            bool wantTable = hint != null && (
                hint.Equals("Table", StringComparison.OrdinalIgnoreCase)
                || hint.Equals("TableThenDesignTab", StringComparison.OrdinalIgnoreCase));
            bool wantChart = hint != null && (
                hint.Equals("Chart", StringComparison.OrdinalIgnoreCase)
                || hint.Equals("ChartThenDesignTab", StringComparison.OrdinalIgnoreCase));

            if (wantTable)
            {
                // 保存比率より、今のセル範囲（B8:F14）を優先する。
                var live = VocabularyHighlightHelper.ResolveHighlights(
                    hwnd, _getExcelApp, hint.Equals("TableThenDesignTab", StringComparison.OrdinalIgnoreCase) ? "Table" : hint);
                if (live.Count > 0)
                    holes.AddRange(live);
                else
                {
                    var cal = ResolveCalibratedHoleNow(table: true);
                    if (cal.HasValue)
                        holes.Add(cal.Value);
                }
                try { VocabularyHighlightHelper.ClearNativeTableHighlight(excel); } catch { }

                if (hint.Equals("TableThenDesignTab", StringComparison.OrdinalIgnoreCase))
                    holes.AddRange(VocabularyHighlightHelper.ResolveHighlights(hwnd, _getExcelApp, "TableDesignTab"));
                return holes;
            }

            if (wantChart)
            {
                var cal = ResolveCalibratedHoleNow(table: false);
                if (cal.HasValue)
                    holes.Add(cal.Value);
                else if (!_calibrationLocked || !HasValidCalibratedRatio(table: false))
                {
                    if (hint.Equals("ChartThenDesignTab", StringComparison.OrdinalIgnoreCase))
                        holes.AddRange(VocabularyHighlightHelper.ResolveHighlights(hwnd, _getExcelApp, "Chart"));
                    else
                        holes.AddRange(VocabularyHighlightHelper.ResolveHighlights(hwnd, _getExcelApp, hint));
                }

                if (hint.Equals("ChartThenDesignTab", StringComparison.OrdinalIgnoreCase))
                    holes.AddRange(VocabularyHighlightHelper.ResolveHighlights(hwnd, _getExcelApp, "ChartDesignTab"));
                try { VocabularyHighlightHelper.ClearNativeTableHighlight(excel); } catch { }
                return holes;
            }

            try { VocabularyHighlightHelper.ClearNativeTableHighlight(excel); } catch { }
            holes.AddRange(VocabularyHighlightHelper.ResolveHighlights(hwnd, _getExcelApp, hint));
            return holes;
        }

        bool ShouldRetryCalibratedHole(string hint, List<Rect> holes)
        {
            if (!_calibrationLocked || string.IsNullOrEmpty(hint)) return false;
            bool table = hint.Equals("Table", StringComparison.OrdinalIgnoreCase)
                         || hint.Equals("TableThenDesignTab", StringComparison.OrdinalIgnoreCase);
            bool chart = hint.Equals("Chart", StringComparison.OrdinalIgnoreCase)
                         || hint.Equals("ChartThenDesignTab", StringComparison.OrdinalIgnoreCase);
            if (table && HasValidCalibratedRatio(true) && !ResolveCalibratedHoleNow(true).HasValue)
                return true;
            if (chart && HasValidCalibratedRatio(false) && !ResolveCalibratedHoleNow(false).HasValue)
                return true;
            if (table && HasValidCalibratedRatio(true)
                && hint.Equals("Table", StringComparison.OrdinalIgnoreCase)
                && holes.Count == 0)
                return true;
            return false;
        }

        /// <summary>表示時に位置が取れなかった案内は、あとから枠を付ける。</summary>
        void ScheduleLiveHoleRetry(string hint)
        {
            var coach = _coach;
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
            int tries = 0;
            timer.Tick += (s, e) =>
            {
                tries++;
                if (_coach == null || _coach != coach || _currentSolved || tries > 12)
                {
                    timer.Stop();
                    return;
                }
                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
                ExcelApp excel = null;
                try { excel = _getExcelApp?.Invoke(); } catch { }
                var holes = BuildCoachHoles(hwnd, excel, hint, null);
                if (holes.Count == 0) return;
                try { _coach.UpdateHighlights(holes); } catch { }
                timer.Stop();
            };
            timer.Start();
        }

        void ScheduleCalibratedHoleRetry(string hint)
        {
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            int tries = 0;
            timer.Tick += (s, e) =>
            {
                tries++;
                if (_coach == null || _currentSolved)
                {
                    timer.Stop();
                    return;
                }

                bool table = hint != null && (
                    hint.Equals("Table", StringComparison.OrdinalIgnoreCase)
                    || hint.Equals("TableThenDesignTab", StringComparison.OrdinalIgnoreCase));
                var cal = ResolveCalibratedHoleNow(table: table);
                if (cal.HasValue)
                {
                    var list = new List<Rect> { cal.Value };
                    if (hint != null && hint.IndexOf("ThenDesignTab", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        IntPtr hwnd = IntPtr.Zero;
                        try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
                        string tabHint = table ? "TableDesignTab" : "ChartDesignTab";
                        list.AddRange(VocabularyHighlightHelper.ResolveHighlights(hwnd, _getExcelApp, tabHint));
                    }
                    try { _coach.UpdateHighlights(list); } catch { }
                    timer.Stop();
                    return;
                }

                if (tries >= 10)
                    timer.Stop();
            };
            timer.Start();
        }

        void CloseCoach()
        {
            try
            {
                if (_coach != null)
                {
                    try { _coach.PhysicalClickCaptured -= OnCalibrationPhysicalClick; } catch { }
                    try { _coach.EndClickCapture(); } catch { }
                    _coach.Close();
                    _coach = null;
                }
            }
            catch { }
            _awaitingDismiss = false;
            _wrongCoachVisible = false;
            try
            {
                var excel = _getExcelApp?.Invoke();
                VocabularyHighlightHelper.ClearNativeTableHighlight(excel);
            }
            catch { }
        }

        void Finish()
        {
            _phase = Phase.Finished;
            CloseCoach();
            StopWatcher();
            WriteVocabModeFlag(false);

            List<VocabularyKeywordItem> mistakes;
            try { mistakes = VocabularyMistakeStore.LoadItems(_category); }
            catch { mistakes = new List<VocabularyKeywordItem>(); }

            string message = mistakes.Count > 0
                ? "お疲れさまでした！\n間違えた問題が " + mistakes.Count + " 問あります。"
                : "お疲れさまでした！単語帳を終了します。";

            ShowCoach(
                title: "単語帳",
                message: message,
                hintOverride: null,
                allowDismiss: true,
                clickThrough: false,
                onDismiss: () => _onFinished?.Invoke(),
                holesOverride: new List<Rect>(),
                appendSelectHint: false,
                coverScreen: true,
                centerBubble: true,
                showOkButton: true);

            if (_coach == null) return;
            _coach.SetPrimaryButtonText("終了");
            if (mistakes.Count > 0)
                _coach.SetSecondaryButton("最後に間違えた問題だけを解く", StartMistakesOnly);
        }

        /// <summary>同じ Excel・同じカテゴリで、記録した間違い問題だけをチュートリアルなしで出す。</summary>
        void StartMistakesOnly()
        {
            _coach = null;
            _awaitingDismiss = false;
            var items = VocabularyMistakeStore.LoadItems(_category);
            if (items.Count == 0)
            {
                _onFinished?.Invoke();
                return;
            }

            StopWatcher();
            VocabularyEventWatcher.ClearEvents();
            _mistakesOnlyMode = true;
            _queue = items.OrderBy(_ => Guid.NewGuid()).ToList();
            _index = 0;
            _tutorialMode = false;
            _phase = Phase.Quiz;
            WriteVocabModeFlag(false);
            EnsureA1ThenBegin();
        }

        public void Cancel()
        {
            CloseCoach();
            StopWatcher();
            WriteVocabModeFlag(false);
            _phase = Phase.Idle;
        }

        public static void WriteVocabModeFlag(bool enabled)
        {
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "mos_excel_vocab_mode.txt");
                File.WriteAllText(path, enabled ? "1" : "0");
            }
            catch { }
        }

        public void Dispose()
        {
            Cancel();
        }
    }
}
