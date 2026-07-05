using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace JARVIS
{
    public partial class EngineerSettingsWindow : Window
    {
        private readonly AppConfig _config;

        public EngineerSettingsWindow(AppConfig config)
        {
            InitializeComponent();
            _config = config;

            // Подставляем то, что уже было сохранено ранее (если было)
            TbEngineerName.Text = _config.EngineerName ?? "";
            TbEngineerPosition.Text = _config.EngineerPosition ?? "";
            TbEngineerIIN.Text = _config.EngineerIIN ?? "";
        }

        private void TbEngineerIIN_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !Regex.IsMatch(e.Text, "^[0-9]+$");
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            _config.EngineerName = TbEngineerName.Text.Trim();
            _config.EngineerPosition = TbEngineerPosition.Text.Trim();
            _config.EngineerIIN = TbEngineerIIN.Text.Trim();
            _config.Save();

            MessageBox.Show("Данные инженера сохранены.", "Готово",
                MessageBoxButton.OK, MessageBoxImage.Information);
            Close();
        }
    }
}