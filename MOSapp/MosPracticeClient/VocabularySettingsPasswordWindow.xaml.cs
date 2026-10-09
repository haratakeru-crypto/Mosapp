using System.Windows;
using System.Windows.Input;

namespace MosPracticeClient
{
    /// <summary>単語帳の位置設定を開く前のパスワード確認。</summary>
    public partial class VocabularySettingsPasswordWindow : Window
    {
        const string Password = "mos123";

        public VocabularySettingsPasswordWindow()
        {
            InitializeComponent();
            Loaded += (_, __) => PasswordInput.Focus();
        }

        void OkButton_Click(object sender, RoutedEventArgs e)
        {
            TryAccept();
        }

        void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        void PasswordInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            TryAccept();
            e.Handled = true;
        }

        void TryAccept()
        {
            if (PasswordInput.Password == Password)
            {
                DialogResult = true;
                return;
            }
            ErrorText.Visibility = Visibility.Visible;
            PasswordInput.Clear();
            PasswordInput.Focus();
        }
    }
}
