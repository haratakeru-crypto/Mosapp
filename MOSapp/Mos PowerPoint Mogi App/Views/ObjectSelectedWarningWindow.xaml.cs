using System.Windows;

namespace MOS_PowerPoint_app.Views
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
