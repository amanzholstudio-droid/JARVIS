using System.ComponentModel;

namespace JARVIS
{
    // Модель одного устройства в списке техники.
    // Реализует INotifyPropertyChanged, чтобы UI (TextBox'ы в ItemsControl)
    // сразу видел изменения и пересчитывал итоговый текст.
    public class Device : INotifyPropertyChanged
    {
        private string _type = "";
        private string _inv = "";
        private string _sn = "";

        public string Type
        {
            get => _type;
            set { _type = value; OnPropertyChanged(nameof(Type)); }
        }

        public string Inv
        {
            get => _inv;
            set { _inv = value; OnPropertyChanged(nameof(Inv)); }
        }

        public string SN
        {
            get => _sn;
            set { _sn = value; OnPropertyChanged(nameof(SN)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}