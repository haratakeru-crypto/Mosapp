using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using QRCoder;

namespace MosPracticeClient
{
    public static class ResultSubmitBinder
    {
        public static async Task BindAsync(TextBlock statusText, Image qrImage, int wrongCount, string subject)
        {
            SetUi(statusText, qrImage, "座席表へ送信中…", false, null);

            var profile = ExamineeStore.Load();
            if (profile == null
                || string.IsNullOrWhiteSpace(profile.UniversityName)
                || string.IsNullOrWhiteSpace(profile.PersonName))
            {
                SetUi(statusText, qrImage, "大学名タブで氏名を登録すると座席表に送れます", false, null);
                return;
            }

            if (ExamineeStore.HasSubmitted(subject))
            {
                SetUi(statusText, qrImage, "座席表に送信しました", false, null);
                return;
            }

            var result = await PracticeSubmitClient.SubmitInitialAsync(wrongCount, subject).ConfigureAwait(false);
            if (result.Success)
            {
                SetUi(statusText, qrImage, "座席表に送信しました", false, null);
                return;
            }

            BitmapImage qr = null;
            if (!string.IsNullOrEmpty(result.QrUrl))
                qr = CreateQrImage(result.QrUrl);

            string status = qr != null
                ? "スマートフォンで読み取り、座席表に送ってください"
                : (string.IsNullOrWhiteSpace(result.Error) ? "座席表への送信に失敗しました" : result.Error);
            SetUi(statusText, qrImage, status, qr != null, qr);
        }

        static void SetUi(TextBlock statusText, Image qrImage, string status, bool showQr, BitmapImage qr)
        {
            var target = (DependencyObject)statusText ?? qrImage;
            if (target == null) return;

            void Apply()
            {
                if (statusText != null) statusText.Text = status;
                if (qrImage == null) return;
                if (qr != null) qrImage.Source = qr;
                qrImage.Visibility = showQr ? Visibility.Visible : Visibility.Collapsed;
            }

            if (target.Dispatcher.CheckAccess())
                Apply();
            else
                target.Dispatcher.Invoke(Apply);
        }

        public static BitmapImage CreateQrImage(string url)
        {
            var generator = new QRCodeGenerator();
            var data = generator.CreateQrCode(url ?? "", QRCodeGenerator.ECCLevel.Q);
            var png = new PngByteQRCode(data);
            byte[] bytes = png.GetGraphic(8);
            var image = new BitmapImage();
            using (var stream = new MemoryStream(bytes))
            {
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.StreamSource = stream;
                image.EndInit();
                image.Freeze();
            }
            return image;
        }
    }
}
