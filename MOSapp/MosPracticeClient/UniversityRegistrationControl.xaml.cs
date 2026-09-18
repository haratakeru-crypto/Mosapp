using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace MosPracticeClient
{
    public partial class UniversityRegistrationControl : UserControl
    {
        public event EventHandler Registered;
        public event EventHandler Deleted;

        List<UniversityLookup> _universities = new List<UniversityLookup>();
        bool _suppressSuggest;
        bool _locked;
        string _pendingClassroom = "";

        public UniversityRegistrationControl()
        {
            InitializeComponent();
            Loaded += UniversityRegistrationControl_Loaded;
        }

        async void UniversityRegistrationControl_Loaded(object sender, RoutedEventArgs e)
        {
            LoadProfileIntoFields();
            _universities = LookupsClient.GetCachedOrEmpty();
            SyncClassroomForUniversity();
            try
            {
                var latest = await LookupsClient.RefreshAsync();
                if (latest != null)
                {
                    _universities = latest;
                    SyncClassroomForUniversity();
                    if (!_locked && UniversitySuggestions.Visibility == Visibility.Visible)
                        ShowUniversitySuggestions(UniversityBox.Text);
                }
            }
            catch
            {
            }
        }

        public void LoadProfileIntoFields()
        {
            var profile = ExamineeStore.Load();
            _suppressSuggest = true;
            UniversityBox.Text = profile?.UniversityName ?? "";
            NameBox.Text = profile?.PersonName ?? "";
            _pendingClassroom = profile?.ClassroomName ?? "";
            _suppressSuggest = false;
            _locked = ExamineeStore.IsRegistered;
            ApplyLockState();
        }

        void UniversityBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressSuggest || _locked) return;
            _pendingClassroom = "";
            ShowUniversitySuggestions(UniversityBox.Text);
        }

        void ShowUniversitySuggestions(string query)
        {
            string q = (query ?? "").Trim();
            if (q.Length < 1)
            {
                UniversitySuggestions.ItemsSource = null;
                UniversitySuggestions.Visibility = Visibility.Collapsed;
                SyncClassroomForUniversity();
                return;
            }

            var matches = LookupsClient.FilterUniversities(_universities, q).Take(12).ToList();
            UniversitySuggestions.ItemsSource = matches.Select(u => u.Name).ToList();
            UniversitySuggestions.Visibility = matches.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            SyncClassroomForUniversity();
        }

        void UniversitySuggestions_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            ApplyUniversitySelection();
        }

        void UniversitySuggestions_PreviewMouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (UniversitySuggestions.SelectedItem == null) return;
            ApplyUniversitySelection();
        }

        void ApplyUniversitySelection()
        {
            if (_locked) return;
            var name = UniversitySuggestions.SelectedItem as string;
            if (string.IsNullOrEmpty(name)) return;
            _suppressSuggest = true;
            UniversityBox.Text = name;
            UniversitySuggestions.Visibility = Visibility.Collapsed;
            _pendingClassroom = "";
            _suppressSuggest = false;
            SyncClassroomForUniversity();
        }

        void SyncClassroomForUniversity()
        {
            var rooms = GetClassrooms(FindUniversity(UniversityBox.Text));
            if (rooms.Count == 0)
            {
                ClassroomPanel.Visibility = Visibility.Collapsed;
                ClassroomCombo.ItemsSource = null;
                ClassroomCombo.SelectedItem = null;
                ClassroomCombo.IsEnabled = false;
                ClassroomHint.Text = "";
                return;
            }

            string keep = MatchRoom(rooms, ClassroomCombo.SelectedItem as string)
                ?? MatchRoom(rooms, _pendingClassroom);

            ClassroomPanel.Visibility = Visibility.Visible;
            ClassroomCombo.ItemsSource = rooms;
            ClassroomCombo.IsEnabled = !_locked && rooms.Count >= 2;
            ClassroomCombo.SelectedItem = keep ?? rooms[0];

            ClassroomHint.Text = rooms.Count == 1
                ? "この大学の教室を自動入力しました。"
                : "教室を一覧から選んでください。直接入力はできません。";
        }

        static List<string> GetClassrooms(UniversityLookup uni)
        {
            if (uni?.Classrooms == null) return new List<string>();
            return uni.Classrooms
                .Where(r => !string.IsNullOrWhiteSpace(r))
                .Select(r => r.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        static string MatchRoom(List<string> rooms, string value)
        {
            if (rooms == null || string.IsNullOrWhiteSpace(value)) return null;
            return rooms.FirstOrDefault(r =>
                string.Equals(r, value.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        UniversityLookup FindUniversity(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            return _universities.FirstOrDefault(u =>
                string.Equals(u.Name, name.Trim(), StringComparison.OrdinalIgnoreCase));
        }

        void ApplyLockState()
        {
            UniversityBox.IsReadOnly = _locked;
            NameBox.IsReadOnly = _locked;
            var fieldBg = _locked
                ? new SolidColorBrush(Color.FromRgb(0xF3, 0xF4, 0xF6))
                : Brushes.White;
            UniversityBox.Background = fieldBg;
            NameBox.Background = fieldBg;
            ClassroomCombo.Background = fieldBg;
            if (_locked)
            {
                UniversitySuggestions.ItemsSource = null;
                UniversitySuggestions.Visibility = Visibility.Collapsed;
            }

            RegisterButton.Content = _locked ? "変更" : "登録";
            EditLockHint.Text = _locked
                ? "内容を直すときは「変更」を押してください。"
                : (ExamineeStore.IsRegistered
                    ? "直し終わったら「登録」を押してください。"
                    : "");
            SyncClassroomForUniversity();
        }

        void RegisterButton_Click(object sender, RoutedEventArgs e)
        {
            if (_locked)
            {
                _locked = false;
                ApplyLockState();
                return;
            }

            string university = (UniversityBox.Text ?? "").Trim();
            string person = (NameBox.Text ?? "").Trim();
            if (string.IsNullOrEmpty(university) || string.IsNullOrEmpty(person))
            {
                StatusText.Text = "大学名とお名前を入力してください。";
                return;
            }

            var rooms = GetClassrooms(FindUniversity(university));
            string classroom = "";
            if (rooms.Count == 1)
            {
                classroom = rooms[0];
            }
            else if (rooms.Count >= 2)
            {
                classroom = MatchRoom(rooms, ClassroomCombo.SelectedItem as string) ?? "";
                if (string.IsNullOrEmpty(classroom))
                {
                    StatusText.Text = "教室を一覧から選んでください。";
                    return;
                }
            }

            var existing = ExamineeStore.Load();
            var profile = new ExamineeProfile
            {
                UniversityName = university,
                PersonName = person,
                ClassroomName = classroom,
                SubmittedSubjects = existing?.SubmittedSubjects ?? new List<string>()
            };
            ExamineeStore.Save(profile);
            PresenceClient.NotifyRegistered(profile);
            StatusText.Text = "";
            _locked = true;
            ApplyLockState();
            Registered?.Invoke(this, EventArgs.Empty);
        }

        void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            var confirm = MessageBox.Show(
                "このPCに保存している大学名・氏名を削除します。よろしいですか？",
                "大学情報を削除",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;

            PresenceClient.NotifyUnregisteredAndWait();
            ExamineeStore.Delete();
            _suppressSuggest = true;
            UniversityBox.Text = "";
            NameBox.Text = "";
            _pendingClassroom = "";
            ClassroomCombo.ItemsSource = null;
            ClassroomCombo.SelectedItem = null;
            _suppressSuggest = false;
            UniversitySuggestions.Visibility = Visibility.Collapsed;
            _locked = false;
            ApplyLockState();
            StatusText.Text = "大学情報を削除しました。";
            Deleted?.Invoke(this, EventArgs.Empty);
        }
    }
}
