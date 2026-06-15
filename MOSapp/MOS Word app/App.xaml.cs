using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Libraries;

namespace MOS_Word_app
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// 起動引数で --openProject &lt;groupId&gt; &lt;projectId&gt; が指定された場合の GroupId。null のときは自動で開かない。
        /// </summary>
        public static int? AutoOpenGroupId { get; private set; }

        /// <summary>
        /// 起動引数で --openProject が指定された場合の ProjectId。null のときは自動で開かない。
        /// </summary>
        public static int? AutoOpenProjectId { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            ParseStartupArgs(e?.Args);

            // MainWindow 表示後に VSTO チェック・準備（UI スレッドをブロックしない）
            Dispatcher.BeginInvoke(new Action(() =>
            {
                CheckVSTOAddInStatus();
                VSTOInstallerHelper.StartBackgroundPrepForExam();
            }), DispatcherPriority.ApplicationIdle);
        }

        private static void ParseStartupArgs(string[] args)
        {
            AutoOpenGroupId = null;
            AutoOpenProjectId = null;
            if (args == null || args.Length < 3) return;
            for (int i = 0; i < args.Length - 2; i++)
            {
                if (args[i] != "--openProject") continue;
                if (!int.TryParse(args[i + 1], out int groupId) || groupId < 1 || groupId > 3) continue;
                if (!int.TryParse(args[i + 2], out int projectId) || projectId < 1 || projectId > 10) continue;
                AutoOpenGroupId = groupId;
                AutoOpenProjectId = projectId;
                return;
            }
        }

        /// <summary>
        /// 自動で開く指定をクリアする（二重に開かないように MainWindow から呼ぶ）。
        /// </summary>
        public static void ClearAutoOpen()
        {
            AutoOpenGroupId = null;
            AutoOpenProjectId = null;
        }

        private void CheckVSTOAddInStatus()
        {
            var status = VSTOInstallerHelper.GetInstallStatus();

            if (!status.IsInstalled)
            {
                // 警告ダイアログを表示
                string message = status.GetInstallationMessage();
                string title = "VSTOアドイン未インストール";

                var owner = Current?.MainWindow;
                string body = message + "\n\nこのまま続行しますか？\n（VSTOが必要なタスクの採点が正しく行われない可能性があります）";
                MessageBoxResult result = owner != null
                    ? MessageBox.Show(owner, body, title, MessageBoxButton.YesNo, MessageBoxImage.Warning)
                    : MessageBox.Show(body, title, MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (result == MessageBoxResult.No)
                {
                    // アプリケーションを終了
                    Shutdown();
                }
            }
        }
    }
}
