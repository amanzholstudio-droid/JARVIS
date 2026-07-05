using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media.Imaging;

namespace JARVIS
{
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();
            LoadLogo();
        }

        // ===== Загружаем логотип, вшитый в exe как Embedded Resource =====
        private void LoadLogo()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string? resourceName = assembly.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith("Logo.jpg", StringComparison.OrdinalIgnoreCase));

                if (resourceName == null) return; // логотип не найден — просто не показываем картинку

                using var stream = assembly.GetManifestResourceStream(resourceName)!;
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.StreamSource = stream;
                bitmap.EndInit();
                bitmap.Freeze();

                LogoImage.Source = bitmap;
            }
            catch
            {
                // Если логотип не загрузился — не критично, просто пустое место вместо картинки
            }
        }

        private void Email_Click(object sender, RoutedEventArgs e)
        {
            var link = (Hyperlink)sender;
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(link.NavigateUri.ToString())
            {
                UseShellExecute = true
            });
        }
    }
}