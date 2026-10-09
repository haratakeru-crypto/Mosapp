using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using MosPracticeClient;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>
    /// 単語帳の位置設定モード。チュートリアルとクイズの各画面を本番と同じ文面で出し、
    /// テキストボックスと枠を手で合わせて保存する。
    /// </summary>
    public sealed partial class VocabularySessionController
    {
        sealed class SettingsSlot
        {
            public string Label;
            public string Note;
            public VocabularyKeywordItem Item;
            /// <summary>Excel 側の準備（タブ切り替え、セル選択など）。</summary>
            public Action Prepare;
            /// <summary>本番と同じ表示を出す。</summary>
            public Action Show;
            /// <summary>枠を保存するキー。null なら枠は保存しない。</summary>
            public string HoleKey;
            public string HoleBasis = CoachHoleOverrideStore.BasisExcel;
            public bool ClosesBackstage;
            /// <summary>情報まで進んだ状態（3段目の案内）で表示する。</summary>
            public bool InfoReached;
        }

        readonly List<SettingsSlot> _settingsSlots = new List<SettingsSlot>();
        int _settingsIndex;
        bool _settingsDirty;
        bool _settingsOperatingExcel;
        VocabularySettingsPanel _settingsPanel;
        VocabularyTutorialOkWindow _tutorialOkDialog;
        DispatcherTimer _settingsEditTimer;

        /// <summary>位置設定モードを始める。判定やイベントの監視は行わない。</summary>
        public void StartSettings()
        {
            StopWatcher();
            WriteVocabModeFlag(false);
            _phase = Phase.Settings;
            _tutorialMode = false;
            _mistakesOnlyMode = false;
            TryLoadPersistedCalibration();
            TryResetSheetToA1();

            BuildSettingsSlots();
            if (_settingsSlots.Count == 0)
            {
                MessageBox.Show("設定できる画面がありません。", "単語帳", MessageBoxButton.OK, MessageBoxImage.Information);
                _phase = Phase.Finished;
                _onFinished?.Invoke();
                return;
            }

            _settingsPanel = new VocabularySettingsPanel();
            _settingsPanel.PreviousClicked += () => MoveSettings(-1);
            _settingsPanel.NextClicked += () => MoveSettings(+1);
            _settingsPanel.SaveClicked += SaveSettingsSlot;
            _settingsPanel.ResetClicked += ResetSettingsSlot;
            _settingsPanel.AddHoleClicked += () => _coach?.AddEditHole();
            _settingsPanel.OperateExcelClicked += ToggleOperateExcel;
            _settingsPanel.ExitClicked += ExitSettings;
            _settingsPanel.Show();

            _settingsIndex = 0;
            ShowSettingsSlot();
        }

        void BuildSettingsSlots()
        {
            _settingsSlots.Clear();
            var tabItems = VocabularyCatalog.Filter(VocabularyCategory.TabButton).ToList();
            var functionItems = VocabularyCatalog.Filter(VocabularyCategory.Function).ToList();

            var table = tabItems.FirstOrDefault(IsTableKeyword);
            if (table != null)
            {
                AddTutorialSlots(table);
            }

            foreach (var item in tabItems)
                AddQuizSlots(item);

            foreach (var item in functionItems)
            {
                var it = item;
                _settingsSlots.Add(new SettingsSlot
                {
                    Label = "関数「" + it.Keyword + "」: 不正解",
                    Item = it,
                    Prepare = () => SelectCellA1(),
                    Show = () => ShowWrongCoach(),
                    HoleKey = CoachHoleOverrideStore.HintKey("FormulaBar")
                });
                _settingsSlots.Add(CorrectSlot(it, "関数「" + it.Keyword + "」: 正解！"));
            }
        }

        void AddTutorialSlots(VocabularyKeywordItem table)
        {
            _settingsSlots.Add(new SettingsSlot
            {
                Label = "チュートリアル: キーワードの表示",
                Item = table,
                Prepare = () => { ActivateHome(); _tutorialSubStep = 0; },
                Show = () => ShowTableTutorialStep(),
                HoleKey = CoachHoleOverrideStore.KeywordAnchorKey,
                HoleBasis = CoachHoleOverrideStore.BasisScreen
            });
            _settingsSlots.Add(new SettingsSlot
            {
                Label = "チュートリアル: あとで見直す",
                Note = "『あとで見直す』ボタンに枠を合わせてください。",
                Item = table,
                Prepare = () => { _tutorialSubStep = ReviewLaterStep; },
                Show = () => ShowTableTutorialStep(),
                HoleKey = CoachHoleOverrideStore.ReviewAnchorKey,
                HoleBasis = CoachHoleOverrideStore.BasisScreen
            });
            _settingsSlots.Add(new SettingsSlot
            {
                Label = "チュートリアル: テーブルを選択",
                Item = table,
                Prepare = () => { ActivateHome(); SelectCellA1(); _tutorialSubStep = 1; },
                Show = () => ShowTableTutorialStep(),
                HoleKey = CoachHoleOverrideStore.HintKey("Table")
            });
            _settingsSlots.Add(new SettingsSlot
            {
                Label = "チュートリアル: テーブルデザインタブ",
                Note = "テーブルのセルを選んだ状態で表示しています。『テーブルデザイン』タブに枠を合わせてください。",
                Item = table,
                Prepare = () => { SelectTableCell(); ActivateHome(); _tutorialSubStep = 2; },
                Show = () => ShowTableTutorialStep(),
                HoleKey = CoachHoleOverrideStore.HintKey("TableDesignTab")
            });
            _settingsSlots.Add(new SettingsSlot
            {
                Label = "チュートリアル: 正解です（OK の画面）",
                Item = table,
                Prepare = () => { ActivateHome(); SelectCellA1(); _tutorialSubStep = 3; },
                Show = () => ShowTableCorrectDialog(),
                HoleKey = CoachHoleOverrideStore.CorrectDialogKey,
                HoleBasis = CoachHoleOverrideStore.BasisScreen
            });
            _settingsSlots.Add(new SettingsSlot
            {
                Label = "チュートリアル: 次へボタン",
                Item = table,
                Prepare = () => { _tutorialSubStep = 4; },
                Show = () => ShowNextButtonCoach(),
                HoleKey = CoachHoleOverrideStore.NextAnchorKey,
                HoleBasis = CoachHoleOverrideStore.BasisScreen
            });
            _settingsSlots.Add(new SettingsSlot
            {
                Label = "チュートリアル: それでは、問題を解いてみましょう",
                Item = table,
                Prepare = () => { _tutorialSubStep = 5; },
                Show = () => ShowLetsSolveBubble(resumeTutorial: false)
            });
        }

        void AddQuizSlots(VocabularyKeywordItem item)
        {
            var it = item;
            string name = "「" + it.Keyword + "」";

            if (UsesSequentialFor(it))
            {
                bool table = IsTableKeyword(it);
                _settingsSlots.Add(new SettingsSlot
                {
                    Label = name + ": 不正解（" + (table ? "テーブルの範囲" : "グラフ") + "を選択）",
                    Item = it,
                    Prepare = () => { ActivateHome(); SelectCellA1(); },
                    Show = () => ShowSelectTargetCoach(),
                    HoleKey = CoachHoleOverrideStore.HintKey(table ? "Table" : "Chart")
                });
                _settingsSlots.Add(new SettingsSlot
                {
                    Label = name + ": 不正解（" + (table ? "テーブルデザイン" : "グラフのデザイン") + "タブ）",
                    Note = (table ? "テーブル" : "グラフ") + "を選んだ状態で表示しています。デザインタブに枠を合わせてください。",
                    Item = it,
                    Prepare = () => { if (table) SelectTableCell(); else SelectChart(); ActivateHome(); },
                    Show = () => ShowDesignTabCoach(asWrong: true),
                    HoleKey = CoachHoleOverrideStore.HintKey(table ? "TableDesignTab" : "ChartDesignTab")
                });
                _settingsSlots.Add(CorrectSlot(it, name + ": 正解！"));
                return;
            }

            if (string.IsNullOrWhiteSpace(it.TargetTab))
            {
                _settingsSlots.Add(CorrectSlot(it, name + ": 正解！"));
                return;
            }

            if (IsFileTargetFor(it))
            {
                _settingsSlots.Add(new SettingsSlot
                {
                    Label = name + ": 不正解（ファイルタブ）",
                    Item = it,
                    Prepare = () => { ActivateHome(); SetWrongState(onTarget: false, backstage: false); },
                    Show = () => ShowWrongCoach(),
                    HoleKey = CoachHoleOverrideStore.QuizKey(it.Keyword, "tab")
                });
                _settingsSlots.Add(new SettingsSlot
                {
                    Label = name + ": 不正解（情報）",
                    Note = "「Excel を操作する」を押してファイルタブを開き、「表示に戻す」を押してから『情報』に枠を合わせてください。",
                    Item = it,
                    Prepare = () => SetWrongState(onTarget: false, backstage: true),
                    Show = () => ShowWrongCoach(),
                    HoleKey = CoachHoleOverrideStore.QuizKey(it.Keyword, "info"),
                    ClosesBackstage = true
                });
                if (!string.IsNullOrWhiteSpace(it.DetailControl))
                {
                    _settingsSlots.Add(new SettingsSlot
                    {
                        Label = name + ": 不正解（" + it.DetailControl + "）",
                        Note = "「Excel を操作する」を押してファイルタブの『情報』を開き、「表示に戻す」を押してから『" + it.DetailControl + "』に枠を合わせてください。",
                        Item = it,
                        Prepare = () => { SetWrongState(onTarget: false, backstage: true); _infoReached = true; },
                        Show = () => ShowWrongCoach(),
                        HoleKey = CoachHoleOverrideStore.QuizKey(it.Keyword, "detail"),
                        ClosesBackstage = true,
                        InfoReached = true
                    });
                }
                _settingsSlots.Add(CorrectSlot(it, name + ": 正解！"));
                return;
            }

            if (IsGroupItem(it))
            {
                _settingsSlots.Add(new SettingsSlot
                {
                    Label = name + ": 不正解（" + it.TargetTab + "タブ）",
                    Item = it,
                    Prepare = () => { ActivateHome(); SetWrongState(onTarget: false, backstage: false); },
                    Show = () => ShowWrongCoach(),
                    HoleKey = CoachHoleOverrideStore.QuizKey(it.Keyword, "tab")
                });
                _settingsSlots.Add(new SettingsSlot
                {
                    Label = name + ": 正解！（『" + it.TargetControl + "』グループを明るく）",
                    Note = "『" + it.TargetTab + "』タブに切り替えて表示しています。『" + it.TargetControl + "』グループに枠を合わせてください。",
                    Item = it,
                    Prepare = () => { SelectTab(it.TargetTab); SetWrongState(onTarget: true, backstage: false); },
                    Show = () => ShowCorrectBubble(),
                    HoleKey = CoachHoleOverrideStore.QuizKey(it.Keyword, "group")
                });
                return;
            }

            _settingsSlots.Add(new SettingsSlot
            {
                Label = name + ": 不正解（" + it.TargetTab + "タブ）",
                Item = it,
                Prepare = () => { ActivateHome(); SetWrongState(onTarget: false, backstage: false); },
                Show = () => ShowWrongCoach(),
                HoleKey = CoachHoleOverrideStore.QuizKey(it.Keyword, "tab")
            });
            if (!string.IsNullOrWhiteSpace(it.TargetControl))
            {
                _settingsSlots.Add(new SettingsSlot
                {
                    Label = name + ": 不正解（" + it.TargetControl + "）",
                    Note = "『" + it.TargetTab + "』タブに切り替えて表示しています。ボタンに枠を合わせてください。",
                    Item = it,
                    Prepare = () => { SelectTab(it.TargetTab); SetWrongState(onTarget: true, backstage: false); },
                    Show = () => ShowWrongCoach(),
                    HoleKey = CoachHoleOverrideStore.QuizKey(it.Keyword, "button")
                });
            }
            _settingsSlots.Add(CorrectSlot(it, name + ": 正解！"));
        }

        SettingsSlot CorrectSlot(VocabularyKeywordItem item, string label)
        {
            return new SettingsSlot
            {
                Label = label,
                Item = item,
                Show = () => ShowCorrectBubble()
            };
        }

        static bool UsesSequentialFor(VocabularyKeywordItem item)
        {
            return item?.DetectKeys != null && item.DetectKeys.Count >= 2
                   && (IsTableKeyword(item) || IsChartKeyword(item));
        }

        static bool IsFileTargetFor(VocabularyKeywordItem item)
        {
            return VocabularyRibbonTabProbe.TabNameEquals(item?.TargetTab, "ファイル");
        }

        static bool IsGroupItem(VocabularyKeywordItem item)
        {
            return item != null
                   && !IsFileTargetFor(item)
                   && string.Equals(item.Kind, "Group", StringComparison.OrdinalIgnoreCase)
                   && !string.IsNullOrWhiteSpace(item.TargetControl);
        }

        void SetWrongState(bool onTarget, bool backstage)
        {
            _onTargetTab = onTarget;
            _backstageOpen = backstage;
        }

        void ShowSettingsSlot()
        {
            if (_phase != Phase.Settings || _settingsSlots.Count == 0) return;
            _settingsIndex = Math.Max(0, Math.Min(_settingsSlots.Count - 1, _settingsIndex));
            var slot = _settingsSlots[_settingsIndex];

            CloseSettingsTransient();
            CloseCoach();
            _settingsDirty = false;
            _settingsOperatingExcel = false;
            _settingsPanel?.SetOperatingExcel(false);

            _current = slot.Item;
            _currentSolved = false;
            _wrongGuideStep = 0;
            _wrongStep = 0;
            _infoReached = false;
            _setKeywordDisplay?.Invoke(_current?.Keyword ?? "");
            _setProgressDisplay?.Invoke((_settingsIndex + 1) + "/" + _settingsSlots.Count);
            _settingsPanel?.SetSlot(_settingsIndex, _settingsSlots.Count, slot.Label, slot.Note, slot.HoleKey != null);

            try { slot.Prepare?.Invoke(); } catch { }
            try { slot.Show?.Invoke(); } catch { }
            BeginEditWhenCoachShown();
        }

        /// <summary>表示が非同期のこともあるので、コーチマークが出たら編集モードにする。</summary>
        void BeginEditWhenCoachShown()
        {
            try { _settingsEditTimer?.Stop(); } catch { }
            int tries = 0;
            CoachMarkOverlayWindow seen = null;
            _settingsEditTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
            _settingsEditTimer.Tick += (s, e) =>
            {
                tries++;
                var coach = _coach;
                // 2回続けて同じウィンドウなら表示が落ち着いたとみなす。
                if (coach != null && coach.IsVisible && coach == seen)
                {
                    _settingsEditTimer.Stop();
                    coach.BeginEditMode();
                    coach.Edited += () => _settingsDirty = true;
                    return;
                }
                seen = coach;
                if (tries > 40) _settingsEditTimer.Stop();
            };
            _settingsEditTimer.Start();
        }

        bool ConfirmLeaveSlot()
        {
            if (!_settingsDirty) return true;
            var answer = MessageBox.Show("この画面の変更を保存しますか？", "単語帳の位置設定",
                MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (answer == MessageBoxResult.Cancel) return false;
            if (answer == MessageBoxResult.Yes) SaveSettingsSlot();
            return true;
        }

        void MoveSettings(int delta)
        {
            if (_phase != Phase.Settings) return;
            int next = _settingsIndex + delta;
            if (next < 0 || next >= _settingsSlots.Count) return;
            if (!ConfirmLeaveSlot()) return;
            LeaveSettingsSlot();
            _settingsIndex = next;
            ShowSettingsSlot();
        }

        void SaveSettingsSlot()
        {
            if (_phase != Phase.Settings || _coach == null) return;
            var slot = _settingsSlots[_settingsIndex];
            try
            {
                _coach.GetEditedBubble(out double left, out double top, out double w, out double h);
                CoachBubblePlacementStore.Confirm(_coach.PlacementMessage, left, top, w, h);

                var holes = _coach.GetEditedHoles();
                if (slot.HoleKey != null && holes.Count > 0)
                    CoachHoleOverrideStore.Set(slot.HoleKey, holes[0], slot.HoleBasis, ExcelHwnd());
                _settingsDirty = false;
                LogQuiz("settings saved " + slot.Label);
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存できませんでした: " + ex.Message, "単語帳の位置設定", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        void ResetSettingsSlot()
        {
            if (_phase != Phase.Settings) return;
            var slot = _settingsSlots[_settingsIndex];
            var answer = MessageBox.Show("この画面のテキストボックスと枠を、自動の位置に戻しますか？", "単語帳の位置設定",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (answer != MessageBoxResult.Yes) return;
            try
            {
                if (_coach != null) CoachBubblePlacementStore.Unconfirm(_coach.PlacementMessage);
                if (slot.HoleKey != null) CoachHoleOverrideStore.Remove(slot.HoleKey);
            }
            catch { }
            ShowSettingsSlot();
        }

        /// <summary>オーバーレイを一時的に隠して、Excel を直接操作できるようにする。</summary>
        void ToggleOperateExcel()
        {
            if (_phase != Phase.Settings) return;
            if (!_settingsOperatingExcel)
            {
                if (!ConfirmLeaveSlot()) return;
                _settingsOperatingExcel = true;
                _settingsPanel?.SetOperatingExcel(true);
                try { _settingsEditTimer?.Stop(); } catch { }
                CloseCoach();
                return;
            }

            // 戻すときは、Excel の状態はそのままで表示だけ出し直す。
            _settingsOperatingExcel = false;
            _settingsPanel?.SetOperatingExcel(false);
            var slot = _settingsSlots[_settingsIndex];
            _current = slot.Item;
            if (slot.ClosesBackstage) SetWrongState(onTarget: false, backstage: true);
            _infoReached = slot.InfoReached;
            _settingsDirty = false;
            try { slot.Show?.Invoke(); } catch { }
            BeginEditWhenCoachShown();
        }

        void LeaveSettingsSlot()
        {
            var slot = _settingsSlots[_settingsIndex];
            if (slot.ClosesBackstage) CloseBackstage(ExcelHwnd());
        }

        void ExitSettings()
        {
            if (_phase != Phase.Settings) return;
            if (!ConfirmLeaveSlot()) return;
            LeaveSettingsSlot();
            CloseSettingsUi();
            CloseCoach();
            _phase = Phase.Finished;
            _onFinished?.Invoke();
        }

        void CloseSettingsTransient()
        {
            var dialog = _tutorialOkDialog;
            _tutorialOkDialog = null;
            if (dialog == null) return;
            // 閉じたときにチュートリアルが次へ進まないよう、段階を外してから閉じる。
            _tutorialSubStep = -1;
            try { dialog.Close(); } catch { }
        }

        void CloseSettingsUi()
        {
            try { _settingsEditTimer?.Stop(); } catch { }
            _settingsEditTimer = null;
            CloseSettingsTransient();
            var panel = _settingsPanel;
            _settingsPanel = null;
            panel?.CloseByCode();
        }

        void ActivateHome()
        {
            try { VocabularyRibbonTabProbe.TryActivateHomeTab(ExcelHwnd()); } catch { }
        }

        void SelectTab(string tab)
        {
            try { VocabularyRibbonTabProbe.TrySelectTab(ExcelHwnd(), tab); } catch { }
        }

        void SelectCellA1()
        {
            try
            {
                var ws = _getExcelApp?.Invoke()?.ActiveSheet as ExcelWorksheet;
                ws?.Range["A1"]?.Select();
            }
            catch { }
        }

        void SelectTableCell()
        {
            try
            {
                ExcelRange range = VocabularyHighlightHelper.TryGetVocabTableRange(_getExcelApp?.Invoke());
                ((ExcelRange)range?.Cells[1, 1])?.Select();
            }
            catch { }
        }

        void SelectChart()
        {
            try
            {
                var ws = _getExcelApp?.Invoke()?.ActiveSheet as ExcelWorksheet;
                dynamic charts = ws?.ChartObjects();
                if (charts != null && charts.Count > 0)
                    charts.Item(1).Activate();
            }
            catch { }
        }
    }
}
