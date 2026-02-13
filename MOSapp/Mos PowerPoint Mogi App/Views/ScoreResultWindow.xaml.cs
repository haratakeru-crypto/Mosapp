using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace MOS_PowerPoint_app.Views
{
    /// <summary>
    /// 採点結果を〇/×で表示するダイアログ
    /// </summary>
    public partial class ScoreResultWindow : Window
    {
        public ScoreResultWindow(IEnumerable<MOS_PowerPoint_app.TaskResult> taskResults)
        {
            InitializeComponent();
            var list = taskResults?.ToList() ?? new List<MOS_PowerPoint_app.TaskResult>();
            DataContext = list;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
