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
        private readonly ScoringRangeProfile _profile;

        public IReadOnlyList<int> SelectedProjectIds { get; private set; }
        public string SelectedRangeLabel { get; private set; }

        public ScoringRangeDialog(
            IReadOnlyCollection<int> availableProjectIds,
            ScoringRangeProfile profile = ScoringRangeProfile.Standard)
        {
            _availableProjectIds = (availableProjectIds ?? Array.Empty<int>())
                .Where(id => id > 0)
                .Distinct()
                .OrderBy(id => id)
                .ToList();
            _profile = profile;
            InitializeComponent();
            ApplyRadioOrder();
            ApplyAvailableRangeOptions();
            BuildCustomCheckboxes();
            if (RangeCustom != null)
                RangeCustom.IsChecked = true;
            UpdateCustomScrollVisibility();
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
            out string rangeLabel,
            ScoringRangeProfile profile = ScoringRangeProfile.Standard)
        {
            selectedIds = Array.Empty<int>();
            rangeLabel = "";
            var dialog = new ScoringRangeDialog(availableProjectIds, profile);
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

        private void ApplyRadioOrder()
        {
            if (RangeOptionsPanel == null)
                return;

            UIElement[] ordered = _profile == ScoringRangeProfile.WordVariant
                ? new UIElement[] { Range1To3, Range4To6, Range7To8, RangeAll, CustomOptionBlock }
                : new UIElement[] { CustomOptionBlock, Range1To7, Range1To8, Range1To9, RangeAll };

            RangeOptionsPanel.Children.Clear();
            foreach (UIElement item in ordered)
            {
                if (item != null)
                    RangeOptionsPanel.Children.Add(item);
            }
        }

        private void ApplyAvailableRangeOptions()
        {
            SetOptionVisibility(Range1To3, ScoringRangeKind.Projects1To3);
            SetOptionVisibility(Range4To6, ScoringRangeKind.Projects4To6);
            SetOptionVisibility(Range7To8, ScoringRangeKind.Projects7To8);
            SetOptionVisibility(Range7To9, ScoringRangeKind.Projects7To9);
            SetOptionVisibility(Range1To7, ScoringRangeKind.Projects1To7);
            SetOptionVisibility(Range1To8, ScoringRangeKind.Projects1To8);
            SetOptionVisibility(Range1To9, ScoringRangeKind.Projects1To9);
            SetOptionVisibility(RangeCustom, ScoringRangeKind.Custom);
            SetOptionVisibility(RangeAll, ScoringRangeKind.All);
        }

        private void SetOptionVisibility(RadioButton button, ScoringRangeKind kind)
        {
            if (button == null)
                return;

            bool offer = ScoringRangeSelection.ShouldOffer(kind, _availableProjectIds, _profile);
            button.Visibility = offer ? Visibility.Visible : Visibility.Collapsed;
            if (!offer && button.IsChecked == true && RangeAll != null)
                RangeAll.IsChecked = true;
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
            UpdateCustomScrollVisibility();
            UpdateStartButtonState();
        }

        private void UpdateCustomScrollVisibility()
        {
            if (CustomScroll == null)
                return;
            CustomScroll.Visibility = RangeCustom != null && RangeCustom.IsChecked == true
                ? Visibility.Visible
                : Visibility.Collapsed;
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
            if (RangeCustom == null)
                return ScoringRangeKind.All;
            if (RangeCustom.IsChecked == true) return ScoringRangeKind.Custom;
            if (Range1To7 != null && Range1To7.IsChecked == true) return ScoringRangeKind.Projects1To7;
            if (Range1To8 != null && Range1To8.IsChecked == true) return ScoringRangeKind.Projects1To8;
            if (Range1To9 != null && Range1To9.IsChecked == true) return ScoringRangeKind.Projects1To9;
            if (Range1To3 != null && Range1To3.IsChecked == true) return ScoringRangeKind.Projects1To3;
            if (Range4To6 != null && Range4To6.IsChecked == true) return ScoringRangeKind.Projects4To6;
            if (Range7To8 != null && Range7To8.IsChecked == true) return ScoringRangeKind.Projects7To8;
            if (Range7To9 != null && Range7To9.IsChecked == true) return ScoringRangeKind.Projects7To9;
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
