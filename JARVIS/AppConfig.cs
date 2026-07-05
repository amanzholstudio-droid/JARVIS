using System.IO;
using System.Text.Json;

namespace JARVIS
{
    // Все настраиваемые пользователем данные приложения: пути к журналу/актам
    // и данные инженера (кто "Сдал"). Хранится в одном JSON-файле в AppData,
    // чтобы не плодить кучу текстовых файлов и легко добавлять новые поля.
    public class AppConfig
    {
        public string? JournalFilePath { get; set; }
        public string? ActsFolderPath { get; set; }
        public string? EngineerName { get; set; }
        public string? EngineerPosition { get; set; }
        public string? EngineerIIN { get; set; }
        public string? MatrixFilePath { get; set; }
        public string? BackupFolderPath { get; set; }
        public DateTime? LastBackupDate { get; set; }
        public string? EmailTo { get; set; }
        public string? EmailCc { get; set; }

        private static readonly string ConfigFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JARVIS");
        private static readonly string ConfigFilePath = Path.Combine(ConfigFolder, "config.json");

        public static AppConfig Load()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    string json = File.ReadAllText(ConfigFilePath);
                    var config = JsonSerializer.Deserialize<AppConfig>(json);
                    if (config != null) return config;
                }
            }
            catch
            {
                // Файл настроек повреждён/нечитаем — просто начинаем с чистого листа,
                // не роняем приложение из-за этого.
            }

            return new AppConfig();
        }

        public void Save()
        {
            Directory.CreateDirectory(ConfigFolder);
            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ConfigFilePath, json);
        }
    }
}