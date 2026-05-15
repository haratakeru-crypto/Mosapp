using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Diagnostics;

namespace PositionExcelApp
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private Process excelProcess = null;

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("shell32.dll")]
        static extern IntPtr SHAppBarMessage(uint dwMessage, ref APPBARDATA pData);

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint ABM_GETTASKBARPOS = 0x00000005;
        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;

        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct APPBARDATA
        {
            public uint cbSize;
            public IntPtr hWnd;
            public uint uCallbackMessage;
            public uint uEdge;
            public RECT rc;
            public IntPtr lParam;
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Excelを起動
            try
            {
                excelProcess = Process.Start("excel.exe");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excelの起動に失敗しました: {ex.Message}",
                              "エラー",
                              MessageBoxButton.OK,
                              MessageBoxImage.Error);
                return;
            }
            
            // Excelのウィンドウが準備できるまで待機
            if (excelProcess != null)
            {
                try
                {
                    excelProcess.WaitForInputIdle(10000); // 最大10秒待機
                    
                    // Excelのメインウィンドウハンドルを取得（リトライロジック）
                    IntPtr excelHwnd = IntPtr.Zero;
                    uint processId = (uint)excelProcess.Id;
                    int retryCount = 0;
                    const int maxRetries = 20; // 最大20回リトライ（10秒）
                    
                    while (excelHwnd == IntPtr.Zero && retryCount < maxRetries)
                    {
                        // プロセスIDからウィンドウハンドルを検索
                        EnumWindows((windowHandle, lParam) =>
                        {
                            GetWindowThreadProcessId(windowHandle, out uint windowProcessId);
                            if (windowProcessId == processId)
                            {
                                // Excelのメインウィンドウを特定（クラス名で判定）
                                StringBuilder className = new StringBuilder(256);
                                GetClassName(windowHandle, className, className.Capacity);
                                if (className.ToString().Contains("XLMAIN"))
                                {
                                    excelHwnd = windowHandle;
                                    return false; // 見つかったので列挙を停止
                                }
                            }
                            return true; // 続行
                        }, IntPtr.Zero);
                        
                        if (excelHwnd == IntPtr.Zero)
                        {
                            Thread.Sleep(500); // 500ms待機してリトライ
                            retryCount++;
                        }
                    }
                    
                    // ウィンドウハンドルが見つかった場合、リサイズ
                    if (excelHwnd != IntPtr.Zero)
                    {
                        // Excelのウィンドウの境界線サイズを取得
                        GetWindowRect(excelHwnd, out RECT excelWindowRect);
                        GetClientRect(excelHwnd, out RECT excelClientRect);
                        
                        int excelBorderWidth = (excelWindowRect.right - excelWindowRect.left) - excelClientRect.right;
                        int excelBorderHeight = (excelWindowRect.bottom - excelWindowRect.top) - excelClientRect.bottom;
                        
                        // Excelのウィンドウを左上 X=0, Y=0、右下 X=1920, Y=774 にリサイズ
                        // 高さ: 258 * 3 = 774 (1032 / 4 * 3)
                        // 境界線を考慮して位置を調整（マージンをゼロにする）
                        int excelX = -excelBorderWidth / 2;
                        int excelY = -excelBorderHeight / 2;
                        int excelWidth = 1920 + excelBorderWidth;
                        int excelHeight = 774 + excelBorderHeight;
                        
                        MoveWindow(excelHwnd, excelX, excelY, excelWidth, excelHeight, true);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Excelウィンドウのリサイズに失敗しました: {ex.Message}",
                                  "エラー",
                                  MessageBoxButton.OK,
                                  MessageBoxImage.Warning);
                }
            }
            
            // ウィンドウハンドルを取得
            IntPtr hWnd = new WindowInteropHelper(this).Handle;
            
            // 現在のウィンドウサイズを取得して境界線のサイズを計算
            GetWindowRect(hWnd, out RECT windowRect);
            GetClientRect(hWnd, out RECT clientRect);
            
            int borderWidth = (windowRect.right - windowRect.left) - clientRect.right;
            int borderHeight = (windowRect.bottom - windowRect.top) - clientRect.bottom;
            
            // ウィンドウを1920x258サイズで、Excelの下に配置
            // 高さ: 258 (1032 / 4)
            // 位置: Y=774 (Excelの下)
            // 境界線を考慮して位置を調整
            int x = -borderWidth / 2; // 左側の境界線を考慮
            int y = 774 - borderHeight / 2; // 上側の境界線を考慮（Excelの下）
            int width = 1920 + borderWidth; // 境界線を含めた幅
            int height = 258 + borderHeight; // 境界線を含めた高さ
            
            MoveWindow(hWnd, x, y, width, height, true);
            
            // 配置後画面に切り替え
            MainPanel.Visibility = Visibility.Collapsed;
            PositionedPanel.Visibility = Visibility.Visible;
        }

        private void ReturnToMainButton_Click(object sender, RoutedEventArgs e)
        {
            // まず、ウィンドウを元のサイズと位置に戻す（即座に実行）
            IntPtr hWnd = new WindowInteropHelper(this).Handle;
            MoveWindow(hWnd, 100, 100, 800, 450, true);
            
            // メイン画面に戻る（即座に実行）
            MainPanel.Visibility = Visibility.Visible;
            PositionedPanel.Visibility = Visibility.Collapsed;
            
            // Excelの保存・終了処理は非同期で実行（UIをブロックしない）
            Task.Run(() =>
            {
                if (excelProcess != null)
                {
                    try
                    {
                        // ExcelのCOMオブジェクトを取得
                        Microsoft.Office.Interop.Excel.Application excelApp = null;
                        try
                        {
                            excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                            
                            // すべてのブックを保存
                            var workbooks = excelApp.Workbooks;
                            foreach (Microsoft.Office.Interop.Excel.Workbook workbook in workbooks)
                            {
                                try
                                {
                                    if (workbook.Path != "")
                                    {
                                        workbook.Save();
                                    }
                                    else
                                    {
                                        // 未保存のブックは保存ダイアログを表示せずに閉じる
                                        workbook.Saved = true;
                                    }
                                    Marshal.ReleaseComObject(workbook);
                                }
                                catch { }
                            }
                            Marshal.ReleaseComObject(workbooks);
                            
                            // Excelを終了
                            excelApp.Quit();
                            Marshal.ReleaseComObject(excelApp);
                            excelApp = null;
                            
                            // プロセスが終了するまで待機（最大5秒）
                            if (!excelProcess.HasExited)
                            {
                                excelProcess.WaitForExit(5000);
                            }
                        }
                        catch (COMException)
                        {
                            // COMオブジェクトが取得できない場合は、プロセスを強制終了
                        }
                        catch (Exception)
                        {
                            // エラーが発生した場合は、プロセスを強制終了
                        }
                        finally
                        {
                            // プロセスがまだ実行中の場合は強制終了
                            try
                            {
                                if (excelProcess != null && !excelProcess.HasExited)
                                {
                                    excelProcess.Kill();
                                    excelProcess.WaitForExit(2000);
                                }
                            }
                            catch { }
                            
                            // プロセスを解放
                            if (excelProcess != null)
                            {
                                try
                                {
                                    excelProcess.Dispose();
                                }
                                catch { }
                                excelProcess = null;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // エラーが発生した場合もプロセスをクリーンアップ
                        try
                        {
                            if (excelProcess != null && !excelProcess.HasExited)
                            {
                                excelProcess.Kill();
                            }
                            excelProcess?.Dispose();
                        }
                        catch { }
                        excelProcess = null;
                    }
                }
            });
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void DetectTaskbarHeight_Click(object sender, RoutedEventArgs e)
        {
            APPBARDATA data = new APPBARDATA();
            data.cbSize = (uint)Marshal.SizeOf(typeof(APPBARDATA));
            
            IntPtr result = SHAppBarMessage(ABM_GETTASKBARPOS, ref data);
            
            if (result != IntPtr.Zero)
            {
                int taskbarHeight = 0;
                int screenWidth = GetSystemMetrics(SM_CXSCREEN);
                int screenHeight = GetSystemMetrics(SM_CYSCREEN);
                
                // タスクバーの位置を判定
                if (data.rc.top == 0 && data.rc.left == 0 && data.rc.right == screenWidth)
                {
                    // 上部
                    taskbarHeight = data.rc.bottom - data.rc.top;
                }
                else if (data.rc.top == 0 && data.rc.left == 0 && data.rc.bottom == screenHeight)
                {
                    // 左側
                    taskbarHeight = data.rc.right - data.rc.left;
                }
                else if (data.rc.left == 0 && data.rc.bottom == screenHeight && data.rc.right == screenWidth)
                {
                    // 下部（一般的なケース）
                    taskbarHeight = data.rc.bottom - data.rc.top;
                }
                else if (data.rc.top == 0 && data.rc.right == screenWidth && data.rc.bottom == screenHeight)
                {
                    // 右側
                    taskbarHeight = data.rc.right - data.rc.left;
                }
                else
                {
                    // その他の場合、高さを計算
                    taskbarHeight = Math.Max(data.rc.bottom - data.rc.top, data.rc.right - data.rc.left);
                }
                
                MessageBox.Show($"タスクバーの縦幅: {taskbarHeight}px\n" +
                              $"位置: Left={data.rc.left}, Top={data.rc.top}, Right={data.rc.right}, Bottom={data.rc.bottom}",
                              "タスクバー情報",
                              MessageBoxButton.OK,
                              MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("タスクバーの情報を取得できませんでした。",
                              "エラー",
                              MessageBoxButton.OK,
                              MessageBoxImage.Error);
            }
        }
    }
}
