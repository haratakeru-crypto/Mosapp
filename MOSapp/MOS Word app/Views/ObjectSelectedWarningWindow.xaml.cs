using System.Windows;

namespace MOS_Word_app.Views
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
