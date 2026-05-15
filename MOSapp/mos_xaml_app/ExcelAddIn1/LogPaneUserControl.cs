using System;
using System.IO;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    /// <summary>
    /// 「ログを表示」でタスクペインに表示するログビュー。
    /// </summary>
    public partial class LogPaneUserControl : UserControl
    {
        public LogPaneUserControl()
        {
            InitializeComponent();
        }

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
