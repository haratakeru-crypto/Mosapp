using System.Windows;
using Libraries;

namespace MOS_PowerPoint_app.Views
{
    public partial class BulkResetScopeDialog : Window
    {
        public PowerPointBulkResetScope SelectedScope { get; private set; }

        public BulkResetScopeDialog()
        {
            InitializeComponent();
            OkButton.IsEnabled = false;
        }

        private void ScopeOption_Changed(object sender, RoutedEventArgs e)
        {
            bool allSelected = ScopeAll.IsChecked == true;
            AllWarningText.Visibility = allSelected ? Visibility.Visible : Visibility.Collapsed;
            OkButton.IsEnabled = ScopeTextbook.IsChecked == true
                || ScopeVariant.IsChecked == true
                || allSelected;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (ScopeTextbook.IsChecked == true)
                SelectedScope = PowerPointBulkResetScope.Textbook;
            else if (ScopeVariant.IsChecked == true)
                SelectedScope = PowerPointBulkResetScope.Variant;
            else if (ScopeAll.IsChecked == true)
                SelectedScope = PowerPointBulkResetScope.All;
            else
                return;

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
