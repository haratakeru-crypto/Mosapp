using System.Windows;
using System.Windows.Controls;

namespace MOS_PowerPoint_app.Views
{
    /// <summary>
    /// PasswordWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class PasswordWindow : Window
    {
        private const string CorrectPassword = "rm04";
        public bool IsPasswordCorrect { get; private set; }

        public PasswordWindow()
        {
            InitializeComponent();
            PasswordBox.Focus();
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            ErrorTextBlock.Visibility = Visibility.Collapsed;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            string input = PasswordBox.Password ?? "";
            if (input == CorrectPassword)
            {
                IsPasswordCorrect = true;
                DialogResult = true;
                Close();
            }
            else
            {
                IsPasswordCorrect = false;
                ErrorTextBlock.Text = "パスワードが正しくありません。";
                ErrorTextBlock.Visibility = Visibility.Visible;
                PasswordBox.Clear();
                PasswordBox.Focus();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            IsPasswordCorrect = false;
            DialogResult = false;
            Close();
        }
    }
}
