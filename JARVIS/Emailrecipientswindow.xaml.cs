using System.Windows;

namespace JARVIS
{
    public partial class EmailRecipientsWindow : Window
    {
        private readonly AppConfig _config;

        public EmailRecipientsWindow(AppConfig config)
        {
            InitializeComponent();
            _config = config;

            TbTo.Text = _config.EmailTo ?? "";
            TbCc.Text = _config.EmailCc ?? "";
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            _config.EmailTo = TbTo.Text.Trim();
            _config.EmailCc = TbCc.Text.Trim();
            _config.Save();

            MessageBox.Show("Получатели сохранены.", "Готово",
                MessageBoxButton.OK, MessageBoxImage.Information);
            Close();
        }
    }
}