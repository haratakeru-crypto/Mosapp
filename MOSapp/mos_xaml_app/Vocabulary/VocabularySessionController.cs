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
        /// <summary>クリック校正中。</summary>
        bool _calibratingHole;
        int _calibrateCornerIndex;
        Point? _calibrateTopLeftPhysical;
        /// <summary>校正比率（物理座標は表示のたびに再計算）。</summary>
        VocabularyHighlightCalibration.HoleRatio _calibratedTableRatio;
        VocabularyHighlightCalibration.HoleRatio _calibratedChartRatio;
        /// <summary>locked 校正があるとき COM 穴に落とさない。</summary>
        bool _calibrationLocked;

        public VocabularySessionController(
            Dispatcher dispatcher,
            Action<string> setKeywordDisplay,
            Action<string> setProgressDisplay,
            Action onFinished,
            Func<IntPtr> getExcelHwnd,
            Func<ExcelApp> getExcelApp = null)
        {
            _dispatcher = dispatcher;
            _setKeywordDisplay = setKeywordDisplay;
            _setProgressDisplay = setProgressDisplay;
            _onFinished = onFinished;
            _getExcelHwnd = getExcelHwnd;
            _getExcelApp = getExcelApp;
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
            var catalog = VocabularyCatalog.Load();
            var quizItems = VocabularyCatalog.Filter(category).OrderBy(_ => Guid.NewGuid()).ToList();

            var tutorial = new List<VocabularyKeywordItem>();
            if (category == VocabularyCategory.TabButton || category == VocabularyCategory.Both)
            {
                var tab = VocabularyCatalog.FindById(catalog.TutorialTabKeywordId)
                          ?? quizItems.FirstOrDefault(i => !i.IsFunction);
                if (tab != null) tutorial.Add(tab);
            }
            if (category == VocabularyCategory.Function || category == VocabularyCategory.Both)
            {
                var fn = VocabularyCatalog.FindById(catalog.TutorialFunctionKeywordId)
                         ?? quizItems.FirstOrDefault(i => i.IsFunction);
                if (fn != null) tutorial.Add(fn);
            }

            _queue = tutorial.Count > 0
                ? tutorial.Concat(quizItems).ToList()
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
            WriteVocabModeFlag(true);
            _calibrationLocked = false;
            _calibratedTableRatio = null;
            _calibratedChartRatio = null;
            TryLoadPersistedCalibration();
            StartWatcher();
            ShowCurrent(showTutorialCoach: _tutorialMode);
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

            _index++;
            if (_index >= _queue.Count)
            {
                Finish();
                return;
            }

            bool stillTutorial = _tutorialMode && _index < CountLeadingTutorial();
            _phase = stillTutorial ? Phase.Tutorial : Phase.Quiz;
            ShowCurrent(showTutorialCoach: stillTutorial);
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
            RebuildAcceptedKeys();
            VocabularyEventWatcher.ClearEvents();
            WriteVocabModeFlag(false);
            WriteVocabModeFlag(true);

            string prefix = _phase == Phase.Tutorial ? "【チュートリアル】" : "";
            _setKeywordDisplay?.Invoke(prefix + _current.Keyword);
            _setProgressDisplay?.Invoke($"{_index + 1}/{_queue.Count}");

            if (showTutorialCoach && NeedsTwoStepTutorial(_current))
            {
                PrepareTutorialStep1();
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

        /// <summary>1/2 開始: 既にテーブル選択済みだと即 2/2 になるため、選択を外す。</summary>
        void PrepareTutorialStep1()
        {
            try
            {
                var excel = _getExcelApp?.Invoke();
                if (excel != null)
                {
                    ExcelApp xl = excel;
                    try
                    {
                        var ws = xl.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                        ws?.Range["A1"]?.Select();
                    }
                    catch { }
                }
            }
            catch { }

            _prevPollTargetSelected = true;
            _step0IgnorePoll = false;
            _suppressSelectionPollUntil = DateTime.UtcNow.AddMilliseconds(800);
            // 抑制明けにベースラインを取り直す
            _dispatcher.BeginInvoke(new Action(() =>
            {
                if (_tutorialSubStep != 0 || _currentSolved) return;
                CapturePollBaseline();
                // まだテーブル上ならポーリングでは進めない（クリック／イベント待ち）
                if (_prevPollTargetSelected)
                    _step0IgnorePoll = true;
            }), DispatcherPriority.ApplicationIdle);
        }

        /// <summary>2/2 開始: デザインタブが既に選択されていても即正解にしない。</summary>
        void PrepareTutorialStep2()
        {
            IntPtr hwnd = IntPtr.Zero;
            try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }
            try { VocabularyRibbonTabProbe.TryActivateHomeTab(hwnd); } catch { }

            _prevPollTargetSelected = true;
            _suppressSelectionPollUntil = DateTime.UtcNow.AddMilliseconds(600);
            _dispatcher.BeginInvoke(new Action(() =>
            {
                if (_tutorialSubStep != 1 || _currentSolved) return;
                CapturePollBaseline();
            }), DispatcherPriority.ApplicationIdle);
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
            if (_tutorialSubStep == 0)
            {
                if (IsTableKeyword(_current))
                    return excel != null && VocabularyHighlightHelper.IsTableCurrentlySelected(excel);
                if (IsChartKeyword(_current))
                    return excel != null && VocabularyHighlightHelper.IsChartCurrentlySelected(excel);
                return false;
            }

            if (_tutorialSubStep == 1)
            {
                if (IsTableKeyword(_current))
                    return VocabularyRibbonTabProbe.IsTableDesignTabSelected(hwnd);
                if (IsChartKeyword(_current))
                    return VocabularyRibbonTabProbe.IsChartDesignTabSelected(hwnd);
            }

            return false;
        }

        void ShowTutorialCoachStep()
        {
            if (_current == null) return;

            // テーブル／グラフ: 1) 対象を選択 → 2) デザインタブをクリック
            if (NeedsTwoStepTutorial(_current))
            {
                if (_tutorialSubStep == 0)
                {
                    bool isTable = IsTableKeyword(_current);
                    ShowCoach(
                        title: "チュートリアル（1/2）",
                        message: isTable
                            ? "ハイライトされたテーブルをクリックして選択してください。"
                            : "ハイライトされたグラフをクリックして選択してください。",
                        hintOverride: isTable ? "Table" : "Chart",
                        allowDismiss: false,
                        clickThrough: true,
                        onDismiss: null);
                    return;
                }

                bool table = IsTableKeyword(_current);
                ShowCoach(
                    title: "チュートリアル（2/2）",
                    message: table
                        ? "リボンの『テーブルデザイン』タブをクリックしてください。"
                        : "リボンの『グラフのデザイン』タブをクリックしてください。",
                    hintOverride: table ? "TableDesignTab" : "ChartDesignTab",
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
                    bool now = IsTutorialTargetCurrentlyMet(excel, hwnd);

                    if (_tutorialSubStep == 0 && _step0IgnorePoll)
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

                    if (_tutorialSubStep == 0)
                    {
                        if (IsTableKeyword(_current))
                            TryAdvanceTutorialBySelection("SelectTable");
                        else if (IsChartKeyword(_current))
                            TryAdvanceTutorialBySelection("SelectChart");
                    }
                    else if (_tutorialSubStep == 1)
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
            }
            catch { }
        }

        void StopWatcher()
        {
            try { _pollTimer?.Stop(); } catch { }
            _pollTimer = null;
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

                if (_phase == Phase.Tutorial && TryAdvanceTutorialBySelection(key))
                    return;

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
                        // テーブル選択後はデザインタブを案内
                        ShowCoach(
                            title: "次の操作",
                            message: IsTableKeyword(_current)
                                ? "リボンの『テーブルデザイン』タブをクリックしてください。"
                                : "リボンの『グラフのデザイン』タブをクリックしてください。",
                            hintOverride: IsTableKeyword(_current) ? "TableDesignTab" : "ChartDesignTab",
                            allowDismiss: false,
                            clickThrough: true,
                            onDismiss: null);
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
                MarkCorrect();
            else if (IsRelevantWrongAttempt(key))
                ShowWrongCoach();
        }

        void ShowWrongCoach()
        {
            string hint = _current.HighlightHint;
            string msg = _current.CoachMessage ?? ("正解は「" + _current.Answer + "」です。");
            if (IsTableKeyword(_current))
            {
                hint = "TableThenDesignTab";
                msg = "まずテーブルを選択すると、「テーブルデザイン」タブが表示されます。";
            }
            else if (IsChartKeyword(_current))
            {
                hint = "ChartThenDesignTab";
                msg = "まずグラフを選択すると、「グラフのデザイン」タブが表示されます。";
            }

            ShowCoach(
                title: "不正解です。",
                message: msg + "\n（正しい場所を選択／操作すると進みます）",
                hintOverride: hint,
                allowDismiss: false,
                clickThrough: true,
                onDismiss: null);
        }

        void MarkCorrect()
        {
            if (_currentSolved) return;
            _currentSolved = true;
            CloseCoach();
            MessageBox.Show("正解です！", "単語帳", MessageBoxButton.OK, MessageBoxImage.Information);
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

            // 1/2: テーブル／グラフ選択
            if (_tutorialSubStep == 0)
            {
                bool selectOk =
                    (IsTableKeyword(_current) && string.Equals(key, "SelectTable", StringComparison.OrdinalIgnoreCase))
                    || (IsChartKeyword(_current) && string.Equals(key, "SelectChart", StringComparison.OrdinalIgnoreCase));

                if (!selectOk) return false;

                _tutorialAdvanceBusy = true;
                _tutorialSubStep = 1;
                _step0IgnorePoll = false;
                CloseCoach();
                PrepareTutorialStep2();
                ShowTutorialCoachStep();
                _tutorialAdvanceBusy = false;
                return true;
            }

            // 2/2: デザインタブ選択（タイマー自動正解はしない）
            if (_tutorialSubStep == 1)
            {
                bool tabOk =
                    (IsTableKeyword(_current) && string.Equals(key, "TableDesignTab", StringComparison.OrdinalIgnoreCase))
                    || (IsChartKeyword(_current) && string.Equals(key, "ChartDesignTab", StringComparison.OrdinalIgnoreCase));

                if (!tabOk) return false;

                _tutorialAdvanceBusy = false;
                _currentSolved = true;
                CloseCoach();
                MessageBox.Show(
                    IsTableKeyword(_current)
                        ? "正解です！\nテーブルを選ぶと「テーブルデザイン」タブが表示されます。"
                        : "正解です！\nグラフを選ぶと「グラフのデザイン」タブが表示されます。",
                    "単語帳",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
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
            if (key.StartsWith("Select", StringComparison.OrdinalIgnoreCase))
                return !_current.IsFunction;
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
            bool beginClickCapture = false)
        {
            try
            {
                CloseCoach();
                _awaitingDismiss = allowDismiss;
                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = _getExcelHwnd?.Invoke() ?? IntPtr.Zero; } catch { }

                string hint = hintOverride ?? _current?.HighlightHint;
                ExcelApp excel = null;
                try { excel = _getExcelApp?.Invoke(); } catch { }

                var holes = BuildCoachHoles(hwnd, excel, hint, holesOverride);

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
                _coach.ShowCoachMark(hwnd, holes, title, message, allowDismiss, clickThrough, appendSelectHint);

                if (beginClickCapture)
                {
                    _coach.PhysicalClickCaptured += OnCalibrationPhysicalClick;
                    _coach.BeginClickCapture();
                }
                else if (ShouldRetryCalibratedHole(hint, holes))
                {
                    ScheduleCalibratedHoleRetry(hint);
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
                var cal = ResolveCalibratedHoleNow(table: true);
                if (cal.HasValue)
                {
                    holes.Add(cal.Value);
                    try { VocabularyHighlightHelper.ClearNativeTableHighlight(excel); } catch { }
                }
                else if (!_calibrationLocked || !HasValidCalibratedRatio(table: true))
                {
                    try { VocabularyHighlightHelper.ApplyNativeTableHighlight(excel); } catch { }
                    if (hint.Equals("TableThenDesignTab", StringComparison.OrdinalIgnoreCase))
                        holes.AddRange(VocabularyHighlightHelper.ResolveHighlights(hwnd, _getExcelApp, "Table"));
                    else
                        holes.AddRange(VocabularyHighlightHelper.ResolveHighlights(hwnd, _getExcelApp, hint));
                }

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
            MessageBox.Show("単語帳を終了します。", "単語帳", MessageBoxButton.OK, MessageBoxImage.Information);
            _onFinished?.Invoke();
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
