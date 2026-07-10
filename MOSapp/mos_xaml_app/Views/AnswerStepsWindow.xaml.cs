using System.Windows;

namespace MOSExcelMogiApp.Views
{
    public partial class AnswerStepsWindow : Window
    {
        public AnswerStepsWindow(string title, string answerSteps)
        {
            InitializeComponent();
            HeaderTextBlock.Text = title;
            AnswerStepsTextBlock.Text = string.IsNullOrWhiteSpace(answerSteps)
                ? "解答手順が登録されていません。"
                : answerSteps;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
