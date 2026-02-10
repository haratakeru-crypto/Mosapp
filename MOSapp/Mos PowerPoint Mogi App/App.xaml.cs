using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace MOS_PowerPoint_app
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // 未処理の例外をキャッチ
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            
            base.OnStartup(e);

            Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose;

            // PowerPoint 画面を直接表示（表紙・パスワードは使用しない）
            var mainWindow = new MainWindow();
            Application.Current.MainWindow = mainWindow;
            mainWindow.Show();
        }
        
        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            string errorMessage = $"未処理の例外が発生しました:\n\n{e.Exception.Message}\n\nスタックトレース:\n{e.Exception.StackTrace}";
            
            if (e.Exception.InnerException != null)
            {
                errorMessage += $"\n\n内部例外:\n{e.Exception.InnerException.Message}";
            }
            
            MessageBox.Show(errorMessage, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            System.Diagnostics.Debug.WriteLine($"DispatcherUnhandledException: {e.Exception}");
            
            // アプリケーションを継続させる（デバッグ用）
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
                {
                    errorMessage += $"\n\n内部例外:\n{ex.InnerException.Message}";
                }
            }
            
            MessageBox.Show(errorMessage, "致命的なエラー", MessageBoxButton.OK, MessageBoxImage.Error);
            System.Diagnostics.Debug.WriteLine($"UnhandledException: {ex}");
        }
    }
}
