using System;
using System.Timers;
using System.Threading.Tasks;
using Microsoft.Maui.Devices;
using Microsoft.Maui.Controls;

namespace MOSExcelMogiApp.Maui.Pages;

public partial class UiTestAppBarPage : ContentPage
{
    private System.Timers.Timer? _timer;
    private TimeSpan _remainingTime;
    private int _currentProjectId = 1;
    private bool _isPaused = false;
    private bool _timerDisabled = false;

    public UiTestAppBarPage()
    {
        InitializeComponent();
        InitializeTimer();
        UpdateProjectTitle();
        // コンストラクタでは呼ばない（Handlerがまだ初期化されていない）
    }
    
    private void SetPageSize()
    {
        var displayInfo = DeviceDisplay.MainDisplayInfo;
        var screenHeight = displayInfo.Height / displayInfo.Density;
        var screenWidth = displayInfo.Width / displayInfo.Density;
        
        // 画面の1/3の高さに設定（下1/3まで表示）
        var targetHeight = screenHeight / 3.0;
        
        // ウィンドウの位置を画面の下2/3の位置に設定（下1/3まで表示）
        var targetY = screenHeight * 2.0 / 3.0;
        
        System.Diagnostics.Debug.WriteLine($"SetPageSize: Screen: {screenWidth}x{screenHeight}, Target: {screenWidth}x{targetHeight} at Y={targetY}");
        
#if WINDOWS
        try
        {
            // Windowを取得（複数の方法を試す）
            var window = GetParentWindow() ?? 
                        (Application.Current?.Windows.Count > 0 ? Application.Current.Windows[0] : null);
            if (window != null)
            {
                // Handlerが初期化されるまで待つ
                if (window.Handler == null)
                {
                    System.Diagnostics.Debug.WriteLine("SetPageSize: Handler is null, retrying...");
                    Task.Delay(500).ContinueWith(_ =>
                    {
                        MainThread.BeginInvokeOnMainThread(() => SetPageSize());
                    });
                    return;
                }
                
                // Windowsプラットフォーム固有のウィンドウ操作
                var platformWindow = window.Handler.PlatformView as Microsoft.UI.Xaml.Window;
                if (platformWindow != null)
                {
                    // ウィンドウのサイズを設定
                    var size = new Windows.Graphics.SizeInt32((int)screenWidth, (int)targetHeight);
                    platformWindow.AppWindow.Resize(size);
                    
                    // ウィンドウの位置を設定（画面の下2/3の位置）
                    var displayArea = Microsoft.UI.Windowing.DisplayArea.GetFromWindowId(
                        platformWindow.AppWindow.Id, 
                        Microsoft.UI.Windowing.DisplayAreaFallback.Nearest);
                    if (displayArea != null)
                    {
                        var x = displayArea.WorkArea.X;
                        var y = (int)targetY;
                        platformWindow.AppWindow.Move(new Windows.Graphics.PointInt32(x, y));
                        
                        System.Diagnostics.Debug.WriteLine($"Windows: Window resized to {screenWidth}x{targetHeight} at position ({x}, {y})");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("SetPageSize: displayArea is null");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"SetPageSize: platformWindow is null. Handler.PlatformView type: {window.Handler?.PlatformView?.GetType().FullName ?? "null"}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("SetPageSize: window is null");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"SetPageSize error: {ex.Message}\n{ex.StackTrace}");
        }
#else
        // その他のプラットフォームではGridの高さを設定
        if (MainGrid != null)
        {
            MainGrid.HeightRequest = targetHeight;
        }
        
        System.Diagnostics.Debug.WriteLine($"Other platform: Screen height: {screenHeight}, Target height: {targetHeight}");
#endif
    }
    
    protected override void OnAppearing()
    {
        base.OnAppearing();
        // ウィンドウが完全に表示された後にサイズを設定（より長い遅延）
        Task.Delay(500).ContinueWith(_ =>
        {
            MainThread.BeginInvokeOnMainThread(() =>
            {
                SetPageSize();
            });
        });
    }
    
    protected override void OnNavigatedTo(NavigatedToEventArgs args)
    {
        base.OnNavigatedTo(args);
        // ナビゲーション後にもサイズを設定
        Task.Delay(300).ContinueWith(_ =>
        {
            MainThread.BeginInvokeOnMainThread(() =>
            {
                SetPageSize();
            });
        });
    }

    private void InitializeTimer()
    {
        // 50分（3000秒）からカウントダウン開始
        _remainingTime = TimeSpan.FromMinutes(50);
        UpdateTimerDisplay();

        // タイマーを1秒間隔で更新
        _timer = new System.Timers.Timer(1000);
        _timer.Elapsed += Timer_Elapsed;
        _timer.AutoReset = true;

        // タイマー無効化フラグを確認（将来的にMainPageから取得）
        _timerDisabled = false; // TODO: MainPage.IsTimerDisabled から取得

        System.Diagnostics.Debug.WriteLine($"UiTestAppBarPage: InitializeTimer called, IsTimerDisabled = {_timerDisabled}");

        if (_timerDisabled)
        {
            // タイマーは開始しない
            _timer.Stop();
            System.Diagnostics.Debug.WriteLine("UiTestAppBarPage: Timer disabled, not starting");
        }
        else
        {
            _timer.Start();
            System.Diagnostics.Debug.WriteLine("UiTestAppBarPage: Timer enabled, starting");
        }

        // 「一時停止」ボタンの状態を更新
        UpdatePauseButtonState();
    }

    private void Timer_Elapsed(object? sender, ElapsedEventArgs e)
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            if (_remainingTime.TotalSeconds > 0)
            {
                _remainingTime = _remainingTime.Subtract(TimeSpan.FromSeconds(1));
                UpdateTimerDisplay();
            }
            else
            {
                _timer?.Stop();
                UpdateTimerDisplay();
            }
        });
    }

    private void UpdateTimerDisplay()
    {
        // タイマーの表示を更新
        if (TimerTextBlock != null)
        {
            TimerTextBlock.Text = _remainingTime.ToString(@"hh\:mm\:ss");
        }
    }

    private void UpdateProjectTitle()
    {
        if (ProjectInfoTextBlock != null)
        {
            ProjectInfoTextBlock.Text = $"プロジェクト {_currentProjectId}/10";
        }
    }

    private void CloseButton_Clicked(object sender, EventArgs e)
    {
        _timer?.Stop();
        // ページを閉じる（ナビゲーションで戻る）
        if (Navigation.NavigationStack.Count > 1)
        {
            Navigation.PopAsync();
        }
    }

    private async void ReviewPageButton_Clicked(object sender, EventArgs e)
    {
        // レビューページに遷移
        var reviewPage = new ReviewPage(
            _remainingTime,
            new Dictionary<int, bool[]>(), // TODO: 実際の完了状態を渡す
            new Dictionary<int, bool[]>(), // TODO: 実際のフラグ状態を渡す
            1 // TODO: 実際のGroupIdを渡す
        );
        
        await Navigation.PushAsync(reviewPage);
    }

    private void ScoreButton_Clicked(object sender, EventArgs e)
    {
        try
        {
            // TODO: 採点機能の実装
            DisplayAlert("情報", "採点機能は準備中です。", "OK");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"採点エラー: {ex.Message}");
            DisplayAlert("エラー", $"採点中にエラーが発生しました: {ex.Message}", "OK");
        }
    }

    private void PauseButton_Clicked(object sender, EventArgs e)
    {
        // タイマーが無効化されている場合は何もしない
        if (_timerDisabled)
        {
            return;
        }

        if (_isPaused)
        {
            // タイマーを再開
            _timer?.Start();
            _isPaused = false;

            // ボタンのテキストを「一時停止」に変更
            if (sender is Button button)
            {
                button.Text = "一時停止";
            }
        }
        else
        {
            // タイマーを停止
            _timer?.Stop();
            _isPaused = true;

            // ボタンのテキストを「再開」に変更
            if (sender is Button button)
            {
                button.Text = "再開";
            }
        }
    }

    private void UpdatePauseButtonState()
    {
        if (PauseButton != null)
        {
            if (_timerDisabled)
            {
                // ボタンを無効化（グレーアウト）
                PauseButton.IsEnabled = false;
                PauseButton.Opacity = 0.5;
                System.Diagnostics.Debug.WriteLine("UiTestAppBarPage: PauseButton disabled (grayed out)");
            }
            else
            {
                // ボタンを有効化
                PauseButton.IsEnabled = true;
                PauseButton.Opacity = 1.0;
                System.Diagnostics.Debug.WriteLine("UiTestAppBarPage: PauseButton enabled");
            }
        }
        else
        {
            System.Diagnostics.Debug.WriteLine("UiTestAppBarPage: PauseButton is null!");
        }
    }

    private void ResetButton_Clicked(object sender, EventArgs e)
    {
        // 確認ダイアログを表示
        DisplayAlert("リセット確認", $"プロジェクト {_currentProjectId} をリセットしますか？", "はい", "いいえ");
        // TODO: 実際のリセット処理を実装
    }

    private void NextProject_Clicked(object sender, EventArgs e)
    {
        // 次のプロジェクトに移動
        if (_currentProjectId < 10)
        {
            _currentProjectId++;
            UpdateProjectTitle();
            DisplayAlert("情報", $"プロジェクト {_currentProjectId} に移動しました。", "OK");
        }
        else
        {
            DisplayAlert("情報", "これが最後のプロジェクトです。", "OK");
        }
    }

    private void PreviousTask_Clicked(object sender, EventArgs e)
    {
        DisplayAlert("情報", "前のタスクに移動します。", "OK");
    }

    private void TaskButton_Clicked(object sender, EventArgs e)
    {
        // タスクボタンがクリックされたときの処理
        if (sender is Button button)
        {
            string taskNumber = button.Text;
            DisplayAlert("情報", $"タスク {taskNumber} を選択しました。", "OK");
        }
    }

    private void NextTask_Clicked(object sender, EventArgs e)
    {
        DisplayAlert("情報", "次のタスクに移動します。", "OK");
    }

    private void CompleteButton_Clicked(object sender, EventArgs e)
    {
        DisplayAlert("情報", "タスクを解答済みにしました。", "OK");
    }

    private void FlagButton_Clicked(object sender, EventArgs e)
    {
        DisplayAlert("情報", "タスクにフラグを設定しました。", "OK");
    }

    private void EndExamButton_Clicked(object sender, EventArgs e)
    {
        // TODO: 試験終了処理を実装
        DisplayAlert("情報", "試験終了機能は準備中です。", "OK");
    }

    protected override void OnDisappearing()
    {
        _timer?.Stop();
        base.OnDisappearing();
    }
}

