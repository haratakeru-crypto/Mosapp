using System.Windows;

namespace MosPracticeClient
{
    public static class ExamStartGuard
    {
        public static bool EnsureRegistered()
        {
            if (ExamineeStore.IsRegistered) return true;
            MessageBox.Show(
                "先に「大学名」タブで大学名とお名前を登録してください。",
                "登録が必要です",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return false;
        }
    }
}
