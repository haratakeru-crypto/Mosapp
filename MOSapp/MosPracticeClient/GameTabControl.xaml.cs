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

        public static readonly DependencyProperty FunctionSnipeUrlProperty =
            DependencyProperty.Register(nameof(FunctionSnipeUrl), typeof(string), typeof(GameTabControl),
                new PropertyMetadata(string.Empty, OnFunctionSnipeUrlChanged));

        public static readonly DependencyProperty ButtonBackgroundProperty =
            DependencyProperty.Register(nameof(ButtonBackground), typeof(Brush), typeof(GameTabControl),
                new PropertyMetadata(new SolidColorBrush(Color.FromRgb(0x42, 0x42, 0x42))));

        public GameTabControl()
        {
            InitializeComponent();
        }

        /// <summary>5秒ジャッジの Web ゲーム URL。</summary>
        public string JudgeUrl
        {
            get => (string)GetValue(JudgeUrlProperty);
            set => SetValue(JudgeUrlProperty, value);
        }

        /// <summary>5秒スナイプの Web ゲーム URL。</summary>
        public string SnipeUrl
        {
            get => (string)GetValue(SnipeUrlProperty);
            set => SetValue(SnipeUrlProperty, value);
        }

        /// <summary>関数スナイプの Web ゲーム URL。空ならボタンを出さない。</summary>
        public string FunctionSnipeUrl
        {
            get => (string)GetValue(FunctionSnipeUrlProperty);
            set => SetValue(FunctionSnipeUrlProperty, value);
        }

        public Brush ButtonBackground
        {
            get => (Brush)GetValue(ButtonBackgroundProperty);
            set => SetValue(ButtonBackgroundProperty, value);
        }

        void JudgeButton_Click(object sender, RoutedEventArgs e) => OpenGame(JudgeUrl, "5秒ジャッジ");

        void SnipeButton_Click(object sender, RoutedEventArgs e) => OpenGame(SnipeUrl, "5秒スナイプ");

        void FunctionSnipeButton_Click(object sender, RoutedEventArgs e) => OpenGame(FunctionSnipeUrl, "関数スナイプ");

        static void OnFunctionSnipeUrlChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (GameTabControl)d;
            if (control.FunctionSnipeButton == null) return;
            control.FunctionSnipeButton.Visibility = string.IsNullOrWhiteSpace(e.NewValue as string)
                ? Visibility.Collapsed
                : Visibility.Visible;
        }

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
