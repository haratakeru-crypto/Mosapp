using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
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

            // #region agent log
            try
            {
                var args = e?.Args ?? new string[0];
                var argsStr = args.Length > 0 ? string.Join("|", args) : "";
                var line1 = "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"App.xaml.cs:OnStartup\",\"message\":\"StartupEventArgs.Args\",\"data\":{\"argsLength\":" + args.Length + ",\"argsJoined\":\"" + argsStr.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"},\"sessionId\":\"debug-session\",\"hypothesisId\":\"B\"}\n";
                var logPath = @"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log";
                try { System.IO.File.AppendAllText(logPath, line1); } catch { System.IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory + "debug.log", line1); }
            }
            catch { }
            // #endregion

            ParseStartupArgs(e?.Args);

            // #region agent log
            try
            {
                var line2 = "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"App.xaml.cs:AfterParse\",\"message\":\"AutoOpen after parse\",\"data\":{\"autoOpenGroupId\":" + (AutoOpenGroupId.HasValue ? AutoOpenGroupId.Value.ToString() : "null") + ",\"autoOpenProjectId\":" + (AutoOpenProjectId.HasValue ? AutoOpenProjectId.Value.ToString() : "null") + "},\"sessionId\":\"debug-session\",\"hypothesisId\":\"B\"}\n";
                var logPath = @"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log";
                try { System.IO.File.AppendAllText(logPath, line2); } catch { System.IO.File.AppendAllText(System.AppDomain.CurrentDomain.BaseDirectory + "debug.log", line2); }
            }
            catch { }
            // #endregion

            // VSTOアドインのインストール状態をチェック
            CheckVSTOAddInStatus();
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

                MessageBoxResult result = MessageBox.Show(
                    message + "\n\nこのまま続行しますか？\n（VSTOが必要なタスクの採点が正しく行われない可能性があります）",
                    title,
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning
                );

                if (result == MessageBoxResult.No)
                {
                    // アプリケーションを終了
                    Shutdown();
                }
            }
        }
    }
}
