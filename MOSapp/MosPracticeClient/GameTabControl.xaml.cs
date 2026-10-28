using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MosPracticeClient
{
    public partial class GameTabControl : UserControl
    {
        public static readonly DependencyProperty JudgeUrlProperty =
            DependencyProperty.Register(nameof(JudgeUrl), typeof(string), typeof(GameTabControl), new PropertyMetadata(string.Empty));

        public static readonly DependencyProperty SnipeUrlProperty =
            DependencyProperty.Register(nameof(SnipeUrl), typeof(string), typeof(GameTabControl), new PropertyMetadata(string.Empty));

        public static readonly DependencyProperty ButtonBackgroundProperty =
            DependencyProperty.Register(nameof(ButtonBackground), typeof(Brush), typeof(GameTabControl),
                new PropertyMetadata(new SolidColorBrush(Color.FromRgb(0x42, 0x42, 0x42))));

        public GameTabControl()
        {
            InitializeComponent();
        }

        /// <summary>3秒ジャッジの Web ゲーム URL。</summary>
        public string JudgeUrl
        {
            get => (string)GetValue(JudgeUrlProperty);
            set => SetValue(JudgeUrlProperty, value);
        }

        /// <summary>3秒スナイプの Web ゲーム URL。</summary>
        public string SnipeUrl
        {
            get => (string)GetValue(SnipeUrlProperty);
            set => SetValue(SnipeUrlProperty, value);
        }

        public Brush ButtonBackground
        {
            get => (Brush)GetValue(ButtonBackgroundProperty);
            set => SetValue(ButtonBackgroundProperty, value);
        }

        void JudgeButton_Click(object sender, RoutedEventArgs e) => OpenGame(JudgeUrl, "3秒ジャッジ");

        void SnipeButton_Click(object sender, RoutedEventArgs e) => OpenGame(SnipeUrl, "3秒スナイプ");

        void OpenGame(string url, string gameName)
        {
            var owner = Window.GetWindow(this);
            if (string.IsNullOrWhiteSpace(url))
            {
                ShowMessage(owner, $"「{gameName}」は準備中です。", MessageBoxImage.Information);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                ShowMessage(owner, $"ブラウザを開けませんでした。\n\n{ex.Message}", MessageBoxImage.Error);
            }
        }

        static void ShowMessage(Window owner, string message, MessageBoxImage image)
        {
            if (owner != null)
                MessageBox.Show(owner, message, "ゲーム", MessageBoxButton.OK, image);
            else
                MessageBox.Show(message, "ゲーム", MessageBoxButton.OK, image);
        }
    }
}
