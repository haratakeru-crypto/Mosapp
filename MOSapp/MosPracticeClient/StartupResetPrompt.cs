using System.Diagnostics;
using System.Windows;

namespace MosPracticeClient
{
    /// <summary>
    /// レンタルPC向け: Release起動時に教材＋類題の一括リセットを確認する。
    /// Debugビルド、およびデバッガ接続時は出さない。
    /// </summary>
    public static class StartupResetPrompt
    {
        public const string Title = "起動時リセット";

        public const string Message =
            "教材（演習）と類題の作業ファイルを初期状態に戻しますか？\n" +
            "前回の編集内容は失われます。";

        public static bool ShouldAsk()
        {
#if DEBUG
            return false;
#else
            return !Debugger.IsAttached;
#endif
        }

        /// <summary>確認を出し、「はい」のときだけ true。</summary>
        public static bool Ask(Window owner = null)
        {
            if (!ShouldAsk())
                return false;

            MessageBoxResult result = owner == null
                ? MessageBox.Show(Message, Title, MessageBoxButton.YesNo, MessageBoxImage.Question)
                : MessageBox.Show(owner, Message, Title, MessageBoxButton.YesNo, MessageBoxImage.Question);

            return result == MessageBoxResult.Yes;
        }
    }
}
