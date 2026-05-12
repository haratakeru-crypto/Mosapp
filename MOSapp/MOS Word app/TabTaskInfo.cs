using System.ComponentModel;

namespace MOS_Word_app
{
    public class TabTaskInfo : INotifyPropertyChanged
    {
        private int _taskNumber;
        private string _question;
        private string _answer;
        private bool _isPassed;

        public int TaskNumber
        {
            get => _taskNumber;
            set
            {
                _taskNumber = value;
                OnPropertyChanged();
            }
        }

        public string Question
        {
            get => _question;
            set
            {
                _question = value;
                OnPropertyChanged();
            }
        }

        public string Answer
        {
            get => _answer;
            set
            {
                _answer = value;
                OnPropertyChanged();
            }
        }

        public bool IsPassed
        {
            get => _isPassed;
            set
            {
                _isPassed = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

