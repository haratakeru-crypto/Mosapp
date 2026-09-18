using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace MosPracticeClient
{
    public partial class ScoringRangeDialog : Window
    {
        private readonly List<int> _availableProjectIds;

        public IReadOnlyList<int> SelectedProjectIds { get; private set; }
        public string SelectedRangeLabel { get; private set; }

        public ScoringRangeDialog(IReadOnlyCollection<int> availableProjectIds)
        {
            _availableProjectIds = (availableProjectIds ?? Array.Empty<int>())
                .Where(id => id > 0)
                .Distinct()
                .OrderBy(id => id)
                .ToList();
            InitializeComponent();
            BuildCustomCheckboxes();
            UpdateStartButtonState();
        }

        public static bool TrySelect(
            Window owner,
            IReadOnlyCollection<int> availableProjectIds,
            out IReadOnlyList<int> selectedIds)
        {
            return TrySelect(owner, availableProjectIds, out selectedIds, out _);
        }

        public static bool TrySelect(
            Window owner,
            IReadOnlyCollection<int> availableProjectIds,
            out IReadOnlyList<int> selectedIds,
            out string rangeLabel)
        {
            selectedIds = Array.Empty<int>();
            rangeLabel = "";
            var dialog = new ScoringRangeDialog(availableProjectIds);
            if (owner != null)
            {
                dialog.Owner = owner;
                dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            else
            {
                dialog.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }

            bool? result = dialog.ShowDialog();
            if (result == true && dialog.SelectedProjectIds != null && dialog.SelectedProjectIds.Count > 0)
            {
                selectedIds = dialog.SelectedProjectIds;
                rangeLabel = dialog.SelectedRangeLabel ?? "";
                return true;
            }

            return false;
        }

        private void BuildCustomCheckboxes()
        {
            CustomCheckPanel.Children.Clear();
            foreach (int id in _availableProjectIds)
            {
                var checkBox = new CheckBox
                {
                    Content = id.ToString(),
                    Tag = id,
                    Margin = new Thickness(0, 0, 12, 8),
                    FontSize = 14,
                    MinWidth = 36
                };
                checkBox.Checked += CustomCheck_Changed;
                checkBox.Unchecked += CustomCheck_Changed;
                CustomCheckPanel.Children.Add(checkBox);
            }
        }

        private void RangeOption_Changed(object sender, RoutedEventArgs e)
        {
            if (CustomScroll == null) return;
            CustomScroll.Visibility = RangeCustom.IsChecked == true
                ? Visibility.Visible
                : Visibility.Collapsed;
            UpdateStartButtonState();
        }

        private void CustomCheck_Changed(object sender, RoutedEventArgs e)
        {
            UpdateStartButtonState();
        }

        private void UpdateStartButtonState()
        {
            if (StartButton == null) return;
            StartButton.IsEnabled = ResolveCurrentSelection().Count > 0;
        }

        private IReadOnlyList<int> ResolveCurrentSelection()
        {
            return ScoringRangeSelection.Resolve(GetSelectedKind(), _availableProjectIds, GetCustomIds());
        }

        private ScoringRangeKind GetSelectedKind()
        {
            if (Range1To3 == null)
                return ScoringRangeKind.All;
            if (Range1To3.IsChecked == true) return ScoringRangeKind.Projects1To3;
            if (Range4To6.IsChecked == true) return ScoringRangeKind.Projects4To6;
            if (Range7To9.IsChecked == true) return ScoringRangeKind.Projects7To9;
            if (Range1To9.IsChecked == true) return ScoringRangeKind.Projects1To9;
            if (RangeCustom.IsChecked == true) return ScoringRangeKind.Custom;
            return ScoringRangeKind.All;
        }

        private IEnumerable<int> GetCustomIds()
        {
            foreach (var child in CustomCheckPanel.Children)
            {
                if (child is CheckBox checkBox && checkBox.IsChecked == true && checkBox.Tag is int id)
                    yield return id;
            }
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            var selected = ResolveCurrentSelection();
            if (selected.Count == 0)
                return;

            SelectedProjectIds = selected;
            SelectedRangeLabel = ScoringRangeSelection.FormatRangeLabel(GetSelectedKind(), selected);
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
