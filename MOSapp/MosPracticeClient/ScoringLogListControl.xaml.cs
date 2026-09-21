using System;
using System.Windows;
using System.Windows.Controls;

namespace MosPracticeClient
{
    public partial class ScoringLogListControl : UserControl
    {
        public static readonly DependencyProperty SubjectProperty = DependencyProperty.Register(
            nameof(Subject),
            typeof(string),
            typeof(ScoringLogListControl),
            new PropertyMetadata("", (d, e) => ((ScoringLogListControl)d).Reload()));

        public event EventHandler<ScoringLogEntry> EntryClicked;

        public string Subject
        {
            get => (string)GetValue(SubjectProperty);
            set => SetValue(SubjectProperty, value);
        }

        public ScoringLogListControl()
        {
            InitializeComponent();
            IsVisibleChanged += ScoringLogListControl_IsVisibleChanged;
            Loaded += (_, __) => Reload();
        }

        void ScoringLogListControl_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (IsVisible)
                Reload();
        }

        public void Reload()
        {
            if (LogList == null) return;
            var entries = string.IsNullOrWhiteSpace(Subject)
                ? new System.Collections.Generic.List<ScoringLogEntry>()
                : ScoringLogStore.Load(Subject);
            LogList.ItemsSource = entries;
            bool empty = entries == null || entries.Count == 0;
            EmptyText.Visibility = empty ? Visibility.Visible : Visibility.Collapsed;
            LogList.Visibility = empty ? Visibility.Collapsed : Visibility.Visible;
        }

        void LogItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is ScoringLogEntry entry)
                EntryClicked?.Invoke(this, entry);
        }
    }
}
