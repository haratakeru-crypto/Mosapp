using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;

namespace MOSExcelMogiApp.Views
{
    /// <summary>
    /// ImageWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class ImageWindow : Window
    {
        public ImageWindow(string imagePath)
        {
            InitializeComponent();
            LoadImage(imagePath);
        }

        private void LoadImage(string imagePath)
        {
            try
            {
                if (File.Exists(imagePath))
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new System.Uri(imagePath, System.UriKind.Absolute);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    
                    AnswerImage.Source = bitmap;
                    
                    // ウィンドウサイズを画像に合わせて調整（最大サイズ制限あり）
                    if (bitmap.PixelWidth > 0 && bitmap.PixelHeight > 0)
                    {
                        double aspectRatio = (double)bitmap.PixelWidth / bitmap.PixelHeight;
                        double maxWidth = 1200;
                        double maxHeight = 800;
                        
                        double displayWidth = bitmap.PixelWidth;
                        double displayHeight = bitmap.PixelHeight;
                        
                        if (displayWidth > maxWidth)
                        {
                            displayWidth = maxWidth;
                            displayHeight = maxWidth / aspectRatio;
                        }
                        
                        if (displayHeight > maxHeight)
                        {
                            displayHeight = maxHeight;
                            displayWidth = maxHeight * aspectRatio;
                        }
                        
                        this.Width = displayWidth + 40; // マージン分を追加
                        this.Height = displayHeight + 100; // ヘッダーとフッター分を追加
                    }
                }
                else
                {
                    MessageBox.Show(
                        $"画像ファイルが見つかりませんでした。\nパス: {imagePath}",
                        "エラー",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error
                    );
                    this.Close();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    $"画像の読み込み中にエラーが発生しました: {ex.Message}",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
                this.Close();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}







