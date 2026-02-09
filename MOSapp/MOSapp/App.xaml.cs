using System;
using System.Windows;

namespace MOSapp
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            base.OnStartup(e);

            Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose;

            var subjectSelectionWindow = new Views.SubjectSelectionWindow();
            Application.Current.MainWindow = subjectSelectionWindow;

            var passwordWindow = new Views.PasswordWindow();
            bool? dialogResult = passwordWindow.ShowDialog();
            if (dialogResult == true && passwordWindow.IsPasswordCorrect)
            {
                subjectSelectionWindow.Show();
            }
            else
            {
                Application.Current.Shutdown();
            }
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            string errorMessage = $"未処理の例外が発生しました:\n\n{e.Exception.Message}\n\nスタックトレース:\n{e.Exception.StackTrace}";
            if (e.Exception.InnerException != null)
                errorMessage += $"\n\n内部例外:\n{e.Exception.InnerException.Message}";
            MessageBox.Show(errorMessage, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            System.Diagnostics.Debug.WriteLine($"DispatcherUnhandledException: {e.Exception}");
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;
            string errorMessage = $"致命的な例外が発生しました:\n\n{(ex != null ? ex.Message : "不明なエラー")}";
            if (ex != null)
            {
                errorMessage += $"\n\nスタックトレース:\n{ex.StackTrace}";
                if (ex.InnerException != null)
                    errorMessage += $"\n\n内部例外:\n{ex.InnerException.Message}";
            }
            MessageBox.Show(errorMessage, "致命的なエラー", MessageBoxButton.OK, MessageBoxImage.Error);
            System.Diagnostics.Debug.WriteLine($"UnhandledException: {ex}");
        }
    }
}
