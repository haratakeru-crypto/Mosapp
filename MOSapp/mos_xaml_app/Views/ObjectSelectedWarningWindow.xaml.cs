using System.Windows;

namespace MOSExcelMogiApp.Views
{
    public partial class ObjectSelectedWarningWindow : Window
    {
        public ObjectSelectedWarningWindow()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
