using System;
using System.IO;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    /// <summary>
    /// デバックタブ「ログを表示」でタスクペインに表示するログ内容用コントロール
    /// </summary>
    public partial class LogPaneUserControl : UserControl
    {
        public LogPaneUserControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// ログファイルを読み込み、テキストボックスに表示する
        /// </summary>
        public void RefreshLog()
        {
            try
            {
                string path = Logger.GetLogFilePath();
                if (File.Exists(path))
                {
                    textBoxLog.Text = File.ReadAllText(path);
                    textBoxLog.SelectionStart = textBoxLog.Text.Length;
                    textBoxLog.ScrollToCaret();
                }
                else
                {
                    textBoxLog.Text = "(ログファイルがありません)";
                }
            }
            catch (Exception ex)
            {
                textBoxLog.Text = "読み込みエラー: " + ex.Message;
            }
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            RefreshLog();
        }
    }
}
