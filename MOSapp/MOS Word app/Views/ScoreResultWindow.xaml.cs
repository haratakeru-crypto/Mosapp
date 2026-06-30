using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace MOS_Word_app.Views
{
    /// <summary>
    /// 採点結果を〇/✖で表示するダイアログ
    /// </summary>
    public partial class ScoreResultWindow : Window
    {
        private DispatcherTimer _keepOnTopTimer;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public ScoreResultWindow(IEnumerable<MOS_Word_app.TaskResult> taskResults)
        {
            InitializeComponent();
            var list = taskResults?.ToList() ?? new List<MOS_Word_app.TaskResult>();
            DataContext = list;

            Loaded += ScoreResultWindow_Loaded;
            Closed += ScoreResultWindow_Closed;
        }

        /// <summary>
        /// 採点結果を最前面のモーダルで表示する（Word が前面に出ても維持）。
        /// </summary>
        public static void ShowResults(Window owner, IEnumerable<MOS_Word_app.TaskResult> taskResults)
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

        private void ScoreResultWindow_Loaded(object sender, RoutedEventArgs e)
        {
            BringToForeground();
            _keepOnTopTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            _keepOnTopTimer.Tick += KeepOnTopTimer_Tick;
            _keepOnTopTimer.Start();
        }

        private void ScoreResultWindow_Closed(object sender, EventArgs e)
        {
            if (_keepOnTopTimer == null)
                return;

            _keepOnTopTimer.Stop();
            _keepOnTopTimer.Tick -= KeepOnTopTimer_Tick;
            _keepOnTopTimer = null;
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
