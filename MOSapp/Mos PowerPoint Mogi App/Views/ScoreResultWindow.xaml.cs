using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace MOS_PowerPoint_app.Views
{
    /// <summary>
    /// 採点結果を〇/×で表示するダイアログ
    /// </summary>
    public partial class ScoreResultWindow : Window
    {
        private DispatcherTimer _keepOnTopTimer;
        private static int _openCount;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public ScoreResultWindow(IEnumerable<MOS_PowerPoint_app.TaskResult> taskResults)
        {
            InitializeComponent();
            var list = taskResults?.ToList() ?? new List<MOS_PowerPoint_app.TaskResult>();
            DataContext = list;

            Loaded += ScoreResultWindow_Loaded;
            Closed += ScoreResultWindow_Closed;
            Deactivated += ScoreResultWindow_Deactivated;
        }

        /// <summary>
        /// 採点結果を最前面のモーダルで表示する（PowerPoint が前面に出ても維持）。
        /// </summary>
        public static void ShowResults(Window owner, IEnumerable<MOS_PowerPoint_app.TaskResult> taskResults)
        {
            var w = new ScoreResultWindow(taskResults)
            {
                Owner = owner,
                Topmost = true,
                ShowInTaskbar = true
            };

            if (owner != null)
            {
                owner.Topmost = true;
                owner.Activate();
            }

            w.ShowDialog();

            if (owner != null)
                owner.Topmost = true;
        }

        /// <summary>
        /// 開いている採点結果ウィンドウを最前面へ戻す（アプリバー操作・Office 配置後などから呼ぶ）。
        /// </summary>
        public static void TryBringOpenToFront()
        {
            try
            {
                var app = Application.Current;
                if (app == null)
                    return;

                void BringAll()
                {
                    foreach (Window window in app.Windows)
                    {
                        if (window is ScoreResultWindow score && score.IsVisible)
                            score.BringToForeground();
                    }
                }

                if (app.Dispatcher.CheckAccess())
                    BringAll();
                else
                    app.Dispatcher.BeginInvoke((Action)BringAll);
            }
            catch { }
        }

        private void ScoreResultWindow_Loaded(object sender, RoutedEventArgs e)
        {
            BringToForeground();
            _keepOnTopTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            _keepOnTopTimer.Tick += KeepOnTopTimer_Tick;
            _keepOnTopTimer.Start();

            if (_openCount++ == 0 && Application.Current != null)
                Application.Current.Activated += Application_Activated;
        }

        private void ScoreResultWindow_Closed(object sender, EventArgs e)
        {
            if (_keepOnTopTimer != null)
            {
                _keepOnTopTimer.Stop();
                _keepOnTopTimer.Tick -= KeepOnTopTimer_Tick;
                _keepOnTopTimer = null;
            }

            if (--_openCount <= 0)
            {
                _openCount = 0;
                if (Application.Current != null)
                    Application.Current.Activated -= Application_Activated;
            }
        }

        private void ScoreResultWindow_Deactivated(object sender, EventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(BringToForeground), DispatcherPriority.ApplicationIdle);
        }

        private static void Application_Activated(object sender, EventArgs e)
        {
            TryBringOpenToFront();
        }

        private void KeepOnTopTimer_Tick(object sender, EventArgs e)
        {
            if (!IsVisible)
                return;

            if (!IsActive)
            {
                Topmost = false;
                Topmost = true;
                BringToForeground();
            }
        }

        private void BringToForeground()
        {
            Topmost = true;
            Activate();
            try
            {
                var helper = new WindowInteropHelper(this);
                if (helper.Handle != IntPtr.Zero)
                    SetForegroundWindow(helper.Handle);
            }
            catch { }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
