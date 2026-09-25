using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MosPracticeClient
{
    public enum VocabularyStartMode
    {
        FromProblem,
        KeywordOnly
    }

    public enum VocabularyCategory
    {
        None,
        TabButton,
        Function,
        Both
    }

    public sealed class VocabularyStartEventArgs : EventArgs
    {
        public VocabularyStartEventArgs(VocabularyStartMode mode, VocabularyCategory category)
        {
            Mode = mode;
            Category = category;
        }

        public VocabularyStartMode Mode { get; }
        public VocabularyCategory Category { get; }
    }

    public partial class VocabularyTabControl : UserControl
    {
        const string FromProblemText = "問題文の中からキーワードを探し、そのキーワードがどこかを探します";
        const string KeywordOnlyText = "キーワードが表示されて、そのキーワードがどこのタブ・ボタン／関数かを探します。下から練習する種類を選んでください。";

        static readonly SolidColorBrush SelectedBackground = Freeze(Color.FromRgb(0xFF, 0x5A, 0x36));
        static readonly SolidColorBrush SelectedForeground = Brushes.White;
        static readonly SolidColorBrush IdleBackground = Freeze(Color.FromRgb(0xE5, 0xE7, 0xEB));
        static readonly SolidColorBrush IdleForeground = Freeze(Color.FromRgb(0x11, 0x18, 0x27));

        enum Mode
        {
            None,
            FromProblem,
            KeywordOnly
        }

        Mode _mode = Mode.None;
        VocabularyCategory _category = VocabularyCategory.None;

        public event EventHandler<VocabularyStartEventArgs> StartRequested;

        public VocabularyTabControl()
        {
            InitializeComponent();
            ApplyModeVisuals();
        }

        void FromProblemButton_Click(object sender, RoutedEventArgs e)
        {
            _mode = Mode.FromProblem;
            _category = VocabularyCategory.None;
            DescriptionText.Text = FromProblemText;
            ApplyModeVisuals();
        }

        void KeywordOnlyButton_Click(object sender, RoutedEventArgs e)
        {
            _mode = Mode.KeywordOnly;
            DescriptionText.Text = KeywordOnlyText;
            ApplyModeVisuals();
        }

        void CategoryTabButton_Click(object sender, RoutedEventArgs e)
        {
            _category = VocabularyCategory.TabButton;
            ApplyModeVisuals();
        }

        void CategoryFunctionButton_Click(object sender, RoutedEventArgs e)
        {
            _category = VocabularyCategory.Function;
            ApplyModeVisuals();
        }

        void CategoryBothButton_Click(object sender, RoutedEventArgs e)
        {
            _category = VocabularyCategory.Both;
            ApplyModeVisuals();
        }

        void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (_mode != Mode.KeywordOnly || _category == VocabularyCategory.None)
                return;

            StartRequested?.Invoke(this, new VocabularyStartEventArgs(
                VocabularyStartMode.KeywordOnly,
                _category));
        }

        static SolidColorBrush Freeze(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        void ApplyModeVisuals()
        {
            bool chosen = _mode != Mode.None;
            DescriptionPanel.Visibility = chosen ? Visibility.Visible : Visibility.Collapsed;

            bool keywordMode = _mode == Mode.KeywordOnly;
            CategoryPanel.Visibility = keywordMode ? Visibility.Visible : Visibility.Collapsed;

            bool canStart = keywordMode && _category != VocabularyCategory.None;
            StartButton.Visibility = canStart ? Visibility.Visible : Visibility.Collapsed;

            bool fromProblem = _mode == Mode.FromProblem;
            FromProblemButton.Background = fromProblem ? SelectedBackground : IdleBackground;
            FromProblemButton.Foreground = fromProblem ? SelectedForeground : IdleForeground;
            FromProblemButton.BorderThickness = new Thickness(0);

            KeywordOnlyButton.Background = keywordMode ? SelectedBackground : IdleBackground;
            KeywordOnlyButton.Foreground = keywordMode ? SelectedForeground : IdleForeground;
            KeywordOnlyButton.BorderThickness = new Thickness(0);

            ApplyCategoryButton(CategoryTabButton, _category == VocabularyCategory.TabButton);
            ApplyCategoryButton(CategoryFunctionButton, _category == VocabularyCategory.Function);
            ApplyCategoryButton(CategoryBothButton, _category == VocabularyCategory.Both);
        }

        static void ApplyCategoryButton(Button button, bool selected)
        {
            button.Background = selected ? SelectedBackground : IdleBackground;
            button.Foreground = selected ? SelectedForeground : IdleForeground;
            button.BorderThickness = new Thickness(0);
        }
    }
}
