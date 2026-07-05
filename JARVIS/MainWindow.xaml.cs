using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Animation;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;

// Явные алиасы: и ClosedXML (через OpenXml.Spreadsheet), и OpenXml.Wordprocessing
// содержат классы с одинаковыми именами Table/TableRow/TableCell/Text.
// Без алиасов компилятор не может понять, какой из них использовать (CS0104).
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace JARVIS
{
    public partial class MainWindow : Window
    {
        // Аналог массива "devices" из JS — ObservableCollection сама уведомляет
        // ItemsControl о добавлении/удалении строк.
        private readonly ObservableCollection<Device> _devices = new();

        // Все сохранённые настройки (пути к журналу/актам, данные инженера)
        private readonly AppConfig _config = AppConfig.Load();

        // Пока false — все обработчики (Checked/TextChanged) игнорируют события.
        // Это нужно, т.к. элементы с заданными в XAML значениями (IsChecked="True",
        // Text="691230400943" и т.п.) генерируют свои события ПРЯМО во время
        // InitializeComponent(), когда часть остальных x:Name-полей ещё не связана —
        // отсюда и NullReferenceException на TbUserName и подобных.
        private bool _isInitialized;

        public MainWindow()
        {
            InitializeComponent();

            _devices.Add(new Device { Type = "монитор", Inv = "000018529", SN = "PHM5P08286" });
            _devices.Add(new Device { Type = "Мини ПК", Inv = "000051808", SN = "FB76XC4" });

            DevicesList.ItemsSource = _devices;

            _isInitialized = true;
            UpdateOutput();

            // Проверяем при каждом запуске: не пора ли делать бэкап журнала
            RunBackupIfDue();
        }

        // ===== Универсальный обработчик — вызывается при любом изменении полей =====
        // Привязан к TextChanged / Checked большинства контролов, чтобы не дублировать
        // логику под каждое событие (как updateUI() вызывался из всех onchange/oninput в HTML).
        private void OnFieldChanged(object sender, RoutedEventArgs e) => UpdateOutput();

        // ===== Кнопка "Добавить" в панели "Оборудование" =====
        private void AddDeviceRow_Click(object sender, RoutedEventArgs e)
        {
            _devices.Add(new Device());
            UpdateOutput();
        }

        // ===== Крестик удаления конкретной карточки устройства =====
        private void DeleteDeviceRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement fe && fe.DataContext is Device device)
            {
                _devices.Remove(device);
                UpdateOutput();
            }
        }

        // ===== Валидация ИИН: разрешаем вводить только цифры (общий для обоих полей ИИН) =====
        private void Iin_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !Regex.IsMatch(e.Text, "^[0-9]+$");
        }

        // ===== Чистка не-цифр при вставке + предупреждение "12 цифр" (общий для TbUserIIN и TbToIIN) =====
        private void IinTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_isInitialized) return;
            if (sender is not TextBox tb) return;

            var digitsOnly = Regex.Replace(tb.Text, @"\D", "");
            if (digitsOnly != tb.Text)
            {
                int caret = tb.CaretIndex;
                tb.Text = digitsOnly;
                tb.CaretIndex = Math.Min(caret, tb.Text.Length);
                return; // TextChanged сработает повторно после присвоения Text
            }

            // Предупреждение "Должно быть 12 цифр" показываем только для основного поля ИИН
            if (tb == TbUserIIN)
            {
                TbIinWarning.Visibility = (tb.Text.Length > 0 && tb.Text.Length < 12)
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }

            UpdateOutput();
        }

        // ===== Переключение вида формы при смене типа операции =====
        private void OperationType_Changed(object sender, RoutedEventArgs e)
        {
            if (!_isInitialized) { UpdateOutput(); return; }

            bool isTransfer = RbTransfer.IsChecked == true;

            // Для "Перемещение" основание (матрица/инцидент) не нужно — инженер тут не участвует
            BasisSection.Visibility = isTransfer ? Visibility.Collapsed : Visibility.Visible;

            // Показываем второй блок полей ("Кому") только для перемещения,
            // и меняем подписи первого блока, чтобы было понятно, что это "От кого"
            ToPersonPanel.Visibility = isTransfer ? Visibility.Visible : Visibility.Collapsed;
            TbEmployeeSectionHeader.Text = isTransfer ? "3. От кого / Кому" : "3. Сотрудник";
            LblUserName.Text = isTransfer ? "ФИО (от кого)" : "ФИО пользователя";
            LblUserIIN.Text = isTransfer ? "ИИН (от кого)" : "ИИН сотрудника";

            UpdateOutput();
        }

        // ===== Основная логика построения итогового текста (аналог updateUI() из JS) =====
        private void UpdateOutput()
        {
            if (!_isInitialized) return; // игнорируем события, всплывшие во время InitializeComponent()

            bool isMatrix = RbMatrix.IsChecked == true;

            // Поле "Номер инцидента" отключается, если основание — по матрице
            TbIncidentNumber.IsEnabled = !isMatrix;

            // Формируем список оборудования (общий для всех типов операции)
            var deviceListBuilder = new StringBuilder();
            if (_devices.Count == 0)
            {
                deviceListBuilder.AppendLine("[Оборудование не добавлено]");
            }
            else
            {
                for (int i = 0; i < _devices.Count; i++)
                {
                    var device = _devices[i];
                    string prefix = _devices.Count > 1 ? $"{i + 1}." : "";
                    string type = string.IsNullOrWhiteSpace(device.Type) ? "[Тип]" : device.Type;
                    string inv = string.IsNullOrWhiteSpace(device.Inv) ? "[Инв. номер]" : device.Inv;
                    string sn = string.IsNullOrWhiteSpace(device.SN) ? "[S/N]" : device.SN;
                    deviceListBuilder.AppendLine($"{prefix}{type} инв. номер {inv} серийный номер {sn}");
                }
            }

            // ===== "Перемещение" — отдельная ветка: два сотрудника, без инженера и без основания =====
            if (RbTransfer.IsChecked == true)
            {
                string fromName = TbUserName.Text.Trim();
                string fromPosition = TbUserPosition.Text.Trim();
                string toName = TbToName.Text.Trim();
                string toPosition = TbToPosition.Text.Trim();

                var transferOutput = new StringBuilder();
                transferOutput.AppendLine(
                    $"Техника перемещена от {(string.IsNullOrWhiteSpace(fromName) ? "[От кого]" : fromName)} " +
                    $"({(string.IsNullOrWhiteSpace(fromPosition) ? "[Должность]" : fromPosition)}) " +
                    $"к {(string.IsNullOrWhiteSpace(toName) ? "[Кому]" : toName)} " +
                    $"({(string.IsNullOrWhiteSpace(toPosition) ? "[Должность]" : toPosition)})");
                transferOutput.Append(deviceListBuilder);
                transferOutput.AppendLine("Примечание: Перемещение");

                TbOutput.Text = transferOutput.ToString();
                return;
            }

            string userName = TbUserName.Text.Trim();
            string userPosition = TbUserPosition.Text.Trim();
            string userIIN = TbUserIIN.Text.Trim();
            string incidentNo = TbIncidentNumber.Text.Trim();

            // Определяем глагол и примечание в зависимости от типа операции
            string verb;
            string note;

            if (RbDelivery.IsChecked == true)
            {
                verb = isMatrix ? "предоставлен" : "выдан";
                note = "Выдача";
            }
            else if (RbReturn.IsChecked == true)
            {
                verb = "сдан";
                note = "Сдача";
            }
            else // Изъятие
            {
                verb = "изъят";
                note = "Изъятие";
            }

            string baseReason = isMatrix ? "матрицей" : "инцидентом";

            var output = new StringBuilder();
            output.AppendLine($"Пользователю {(string.IsNullOrWhiteSpace(userName) ? "[ФИО]" : userName)} " +
                               $"({(string.IsNullOrWhiteSpace(userPosition) ? "[Должность]" : userPosition)}) " +
                               $"в соответствии с {baseReason} {verb} ");
            output.Append(deviceListBuilder);
            output.AppendLine($"Примечание: {note}");

            if (!isMatrix)
            {
                output.AppendLine($"Инцидент № {(string.IsNullOrWhiteSpace(incidentNo) ? "[Номер]" : incidentNo)}");
            }

            output.Append($"ИИН: {(string.IsNullOrWhiteSpace(userIIN) ? "[12 цифр]" : userIIN)}");

            TbOutput.Text = output.ToString();
        }

        // ===== Копирование текста в буфер обмена + сохранение в журнал + показ тоста =====
        private void CopyToClipboard_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveToJournal();
            }
            catch (IOException)
            {
                // Файл, скорее всего, открыт в Excel — предупреждаем и не роняем программу.
                MessageBox.Show(
                    "Не удалось сохранить запись в журнал: файл сейчас открыт в Excel (или занят другой программой).\n" +
                    "Закройте файл и нажмите кнопку ещё раз.",
                    "Журнал занят", MessageBoxButton.OK, MessageBoxImage.Warning);
                return; // копировать в буфер не будем, пока запись в журнал не удалась
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения в журнал: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                Clipboard.SetText(TbOutput.Text);
                ShowToast();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка копирования: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ===== Путь к файлу журнала: сохранённый пользователем, либо путь по умолчанию =====
        private string GetJournalPath()
        {
            if (!string.IsNullOrWhiteSpace(_config.JournalFilePath))
                return _config.JournalFilePath!;

            // Путь по умолчанию: Мои документы\journal.xlsx — доступен на запись
            // любому пользователю на любом доменном ноуте без прав администратора.
            string defaultPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "journal.xlsx");

            SetJournalPath(defaultPath);
            return defaultPath;
        }

        private void SetJournalPath(string path)
        {
            _config.JournalFilePath = path;
            _config.Save();
        }

        // ===== Папка для актов: сохранённая пользователем, либо путь по умолчанию =====
        private string GetActsFolder()
        {
            if (!string.IsNullOrWhiteSpace(_config.ActsFolderPath))
                return _config.ActsFolderPath!;

            string defaultFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Акты");

            SetActsFolder(defaultFolder);
            return defaultFolder;
        }

        private void SetActsFolder(string folder)
        {
            _config.ActsFolderPath = folder;
            _config.Save();
        }

        // ===== Заголовки колонок журнала (используются и при создании файла) =====
        private static readonly string[] JournalHeaders =
        {
            "Дата и время", "Тип операции", "Основание", "№ инцидента",
            "ФИО сотрудника", "Должность", "ИИН",
            "Кому ФИО", "Кому Должность", "Кому ИИН",
            "Тип устройства", "Инв. номер", "Серийный номер"
        };

        // ===== Дозапись строк в xlsx-журнал: одна строка на каждое устройство =====
        private void SaveToJournal()
        {
            string path = GetJournalPath();

            using var workbook = File.Exists(path) ? new XLWorkbook(path) : new XLWorkbook();
            var sheet = workbook.Worksheets.FirstOrDefault() ?? workbook.Worksheets.Add("Журнал");

            // Если лист пустой — пишем заголовки первой строкой
            if (sheet.LastRowUsed() == null)
            {
                for (int col = 0; col < JournalHeaders.Length; col++)
                {
                    sheet.Cell(1, col + 1).Value = JournalHeaders[col];
                    sheet.Cell(1, col + 1).Style.Font.Bold = true;
                }
            }

            int nextRow = (sheet.LastRowUsed()?.RowNumber() ?? 1) + 1;

            bool isTransfer = RbTransfer.IsChecked == true;

            string opType = isTransfer ? "Перемещение"
                          : RbDelivery.IsChecked == true ? "Выдача"
                          : RbReturn.IsChecked == true ? "Сдача" : "Изъятие";
            string basis = isTransfer ? "" : (RbMatrix.IsChecked == true ? "Матрица" : "Инцидент");
            string incidentNo = (isTransfer || RbMatrix.IsChecked == true) ? "" : TbIncidentNumber.Text.Trim();
            string timestamp = DateTime.Now.ToString("dd.MM.yyyy HH:mm");

            string toName = isTransfer ? TbToName.Text.Trim() : "";
            string toPosition = isTransfer ? TbToPosition.Text.Trim() : "";
            string toIIN = isTransfer ? TbToIIN.Text.Trim() : "";

            // Если устройств нет — всё равно фиксируем операцию одной строкой без техники
            var rows = _devices.Count > 0 ? _devices : new ObservableCollection<Device> { new Device() };

            foreach (var device in rows)
            {
                sheet.Cell(nextRow, 1).Value = timestamp;
                sheet.Cell(nextRow, 2).Value = opType;
                sheet.Cell(nextRow, 3).Value = basis;
                sheet.Cell(nextRow, 4).Value = incidentNo;
                sheet.Cell(nextRow, 5).Value = TbUserName.Text.Trim();
                sheet.Cell(nextRow, 6).Value = TbUserPosition.Text.Trim();
                sheet.Cell(nextRow, 7).Value = TbUserIIN.Text.Trim();
                sheet.Cell(nextRow, 8).Value = toName;
                sheet.Cell(nextRow, 9).Value = toPosition;
                sheet.Cell(nextRow, 10).Value = toIIN;
                sheet.Cell(nextRow, 11).Value = device.Type;
                sheet.Cell(nextRow, 12).Value = device.Inv;
                sheet.Cell(nextRow, 13).Value = device.SN;
                nextRow++;
            }

            sheet.Columns().AdjustToContents();
            workbook.SaveAs(path);
        }

        // ===== Раз в 3 дня копируем журнал в сетевую папку, если она настроена =====
        private void RunBackupIfDue()
        {
            if (string.IsNullOrWhiteSpace(_config.BackupFolderPath)) return; // бэкап не настроен — ничего не делаем

            bool due = _config.LastBackupDate == null
                || (DateTime.Now - _config.LastBackupDate.Value) >= TimeSpan.FromDays(3);

            if (!due) return;

            try
            {
                DoBackup();
                _config.LastBackupDate = DateTime.Now;
                _config.Save();
            }
            catch
            {
                // Сетевая папка недоступна (нет VPN, диск не подключен и т.п.) — тихо пропускаем.
                // LastBackupDate НЕ обновляем, чтобы попытаться снова при следующем запуске.
            }
        }

        // ===== Копирование текущего журнала в папку бэкапа с меткой даты/времени =====
        private void DoBackup()
        {
            string journalPath = GetJournalPath();
            if (!File.Exists(journalPath)) return; // журнала ещё нет — копировать нечего

            Directory.CreateDirectory(_config.BackupFolderPath!);
            string backupFileName = $"journal_backup_{DateTime.Now:yyyy-MM-dd_HH-mm}.xlsx";
            File.Copy(journalPath, Path.Combine(_config.BackupFolderPath!, backupFileName), overwrite: true);
        }

        // ===== Пункт меню: указать сетевую папку для бэкапа =====
        private void ChangeBackupFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog
            {
                Title = "Выберите папку для бэкапа журнала (например, сетевой диск)",
                InitialDirectory = string.IsNullOrWhiteSpace(_config.BackupFolderPath)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : _config.BackupFolderPath
            };

            if (dialog.ShowDialog() == true)
            {
                _config.BackupFolderPath = dialog.FolderName;
                _config.Save();

                MessageBox.Show($"Бэкап журнала теперь будет копироваться сюда (раз в 3 дня):\n{dialog.FolderName}",
                    "Папка для бэкапа сохранена", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // ===== Пункт меню: сделать бэкап прямо сейчас, не дожидаясь 3 дней =====
        private void BackupNow_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_config.BackupFolderPath))
            {
                MessageBox.Show("Сначала укажи папку для бэкапа через пункт меню \"Изменить папку для бэкапа...\".",
                    "Папка не настроена", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                DoBackup();
                _config.LastBackupDate = DateTime.Now;
                _config.Save();

                MessageBox.Show("Бэкап журнала успешно создан.", "Готово",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось создать бэкап: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ===== Пункт меню: сменить файл журнала =====
        private void ChangeJournalPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel файл (*.xlsx)|*.xlsx",
                FileName = "journal.xlsx",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                OverwritePrompt = false // выбираем существующий или новый файл, не обязательно перезаписывать
            };

            if (dialog.ShowDialog() == true)
            {
                SetJournalPath(dialog.FileName);

                MessageBox.Show($"Журнал теперь сохраняется сюда:\n{dialog.FileName}",
                    "Путь изменён", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // ===== Пункт меню: открыть папку с текущим файлом журнала в проводнике =====
        private void OpenJournalFolder_Click(object sender, RoutedEventArgs e)
        {
            string path = GetJournalPath();
            string? folder = Path.GetDirectoryName(path);

            if (folder != null && Directory.Exists(folder))
            {
                System.Diagnostics.Process.Start("explorer.exe", folder);
            }
        }

        // ===== Пункт меню: настроить получателей письма =====
        private void OpenEmailRecipients_Click(object sender, RoutedEventArgs e)
        {
            var window = new EmailRecipientsWindow(_config) { Owner = this };
            window.ShowDialog();
        }

        // ===== Кнопка "Отправить по почте" =====
        private void SendEmailWithAct_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_config.EmailTo))
            {
                MessageBox.Show("Сначала настрой получателей через меню \"Настроить получателей письма\".",
                    "Получатели не настроены", MessageBoxButton.OK, MessageBoxImage.Information);
                OpenEmailRecipients_Click(sender, e);
                return;
            }

            bool isTransfer = RbTransfer.IsChecked == true;
            if (!isTransfer && (string.IsNullOrWhiteSpace(_config.EngineerName) || string.IsNullOrWhiteSpace(_config.EngineerIIN)))
            {
                MessageBox.Show("Сначала заполните свои данные через пункт меню \"Мои данные (инженер)\".",
                    "Данные инженера не заполнены", MessageBoxButton.OK, MessageBoxImage.Information);
                OpenEngineerSettings_Click(sender, e);
                return;
            }

            try
            {
                // Сначала генерируем сам акт — прикладываем именно готовый файл, а не текст
                string actPath = GenerateActDocument();

                // Позднее связывание COM вместо Microsoft.Office.Interop.Outlook —
                // не зависит от конкретной версии Office PIA, установленной на компьютере.
                Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                    throw new InvalidOperationException("Outlook не найден на этом компьютере (не установлен или не зарегистрирован).");

                dynamic outlookApp = Activator.CreateInstance(outlookType)!;
                dynamic mail = outlookApp.CreateItem(0); // 0 = olMailItem

                mail.To = _config.EmailTo;
                if (!string.IsNullOrWhiteSpace(_config.EmailCc))
                    mail.CC = _config.EmailCc;

                mail.Subject = $"Акт приёма-передачи — {TbUserName.Text.Trim()} — {DateTime.Now:dd.MM.yyyy}";
                mail.Body = TbOutput.Text;
                mail.Attachments.Add(actPath);

                mail.Display(false); // открываем письмо на экране — отправляет сам инженер, проверив всё глазами
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось создать письмо: {ex.Message}\n\nПроверь, что Outlook установлен и запущен на этом компьютере.",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ===== Пункт меню: сменить папку для актов =====
        private void ChangeActsFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog
            {
                Title = "Выберите папку для сохранения актов",
                InitialDirectory = GetActsFolder()
            };

            if (dialog.ShowDialog() == true)
            {
                SetActsFolder(dialog.FolderName);
                MessageBox.Show($"Акты теперь сохраняются сюда:\n{dialog.FolderName}",
                    "Папка изменена", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // ===== Пункт меню: открыть полезную ссылку (внутренний регламент коллеги) =====
        private void OpenUsefulLink_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(
                "https://rhetorical-pancake-926.notion.site/23f2ccd1d52380728eb1e6fe7f28401c")
            {
                UseShellExecute = true
            });
        }

        // ===== Пункт меню: окно "Об авторе" =====
        private void OpenAbout_Click(object sender, RoutedEventArgs e)
        {
            var window = new AboutWindow { Owner = this };
            window.ShowDialog();
        }

        // ===== Пункт меню: открыть окно поиска техники по журналу =====
        private void OpenJournalSearch_Click(object sender, RoutedEventArgs e)
        {
            var window = new JournalSearchWindow(GetJournalPath()) { Owner = this };
            window.ShowDialog();
        }

        // ===== Пункт меню: открыть окно проверки по матрице =====
        private void OpenMatrixCheck_Click(object sender, RoutedEventArgs e)
        {
            var window = new MatrixCheckWindow(_config.MatrixFilePath) { Owner = this };
            window.ShowDialog();
        }

        // ===== Пункт меню: загрузить/сменить файл матрицы =====
        private void LoadMatrix_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Выберите файл матрицы оборудования",
                Filter = "Excel файл (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    // Пробуем сразу прочитать файл, чтобы не сохранить путь на "битый" файл
                    var entries = MatrixLoader.Load(dialog.FileName);

                    _config.MatrixFilePath = dialog.FileName;
                    _config.Save();

                    MessageBox.Show($"Матрица загружена: {entries.Count} строк.",
                        "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось прочитать файл: {ex.Message}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        // ===== Отдельный пункт меню: "Мои данные (инженер)" =====
        private void OpenEngineerSettings_Click(object sender, RoutedEventArgs e)
        {
            var window = new EngineerSettingsWindow(_config) { Owner = this };
            window.ShowDialog();
        }

        // ===== Кнопка "Скачать акт" =====
        private void GenerateAct_Click(object sender, RoutedEventArgs e)
        {
            bool isTransfer = RbTransfer.IsChecked == true;

            if (!isTransfer && (string.IsNullOrWhiteSpace(_config.EngineerName) || string.IsNullOrWhiteSpace(_config.EngineerIIN)))
            {
                MessageBox.Show("Сначала заполните свои данные через пункт меню \"Мои данные (инженер)\".",
                    "Данные инженера не заполнены", MessageBoxButton.OK, MessageBoxImage.Information);
                OpenEngineerSettings_Click(sender, e);
                return;
            }

            try
            {
                string outputPath = GenerateActDocument();

                // Открываем сразу в Word, чтобы можно было распечатать/подписать
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось создать акт: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ===== Основная генерация акта на основе шаблона =====
        private string GenerateActDocument()
        {
            string folder = GetActsFolder();
            Directory.CreateDirectory(folder);

            string employeeName = TbUserName.Text.Trim();
            bool isTransfer = RbTransfer.IsChecked == true;

            // ===== Кто "Сдал", а кто "Принял" зависит от типа операции =====
            // Выдача: инженер отдаёт технику сотруднику -> Сдал = инженер, Принял = сотрудник
            // Сдача / Изъятие: техника идёт обратно -> Сдал = сотрудник, Принял = инженер
            // Перемещение: между двумя сотрудниками, инженер не участвует
            bool isIssue = RbDelivery.IsChecked == true;

            string handOverName, handOverIIN, handOverPosition;
            string receiveName, receiveIIN, receivePosition;

            if (isTransfer)
            {
                handOverName = employeeName;
                handOverIIN = TbUserIIN.Text.Trim();
                handOverPosition = TbUserPosition.Text.Trim();

                receiveName = TbToName.Text.Trim();
                receiveIIN = TbToIIN.Text.Trim();
                receivePosition = TbToPosition.Text.Trim();
            }
            else if (isIssue)
            {
                handOverName = _config.EngineerName!;
                handOverIIN = _config.EngineerIIN!;
                handOverPosition = _config.EngineerPosition ?? "";

                receiveName = employeeName;
                receiveIIN = TbUserIIN.Text.Trim();
                receivePosition = TbUserPosition.Text.Trim();
            }
            else
            {
                handOverName = employeeName;
                handOverIIN = TbUserIIN.Text.Trim();
                handOverPosition = TbUserPosition.Text.Trim();

                receiveName = _config.EngineerName!;
                receiveIIN = _config.EngineerIIN!;
                receivePosition = _config.EngineerPosition ?? "";
            }

            // ===== Имя файла: <инв>_<тип операции>_<серийный> для каждого устройства =====
            string opTypeForFile = isTransfer ? "перемещение"
                                  : RbDelivery.IsChecked == true ? "выдача"
                                  : RbReturn.IsChecked == true ? "сдача" : "изъятие";

            var devicesForFile = _devices.Count > 0 ? _devices : new ObservableCollection<Device> { new Device() };
            string namePart = string.Join("_", devicesForFile.Select(d =>
                $"{SanitizeForFileName(d.Inv)}_{opTypeForFile}_{SanitizeForFileName(d.SN)}"));

            string fileName = $"{namePart}_{DateTime.Now:dd.MM.yyyy_HH-mm}.docx";
            string outputPath = Path.Combine(folder, fileName);

            byte[] templateBytes = LoadTemplateBytes();
            File.WriteAllBytes(outputPath, templateBytes);

            using (var doc = WordprocessingDocument.Open(outputPath, true))
            {
                var body = doc.MainDocumentPart!.Document.Body!;

                // ===== Дата =====
                string newDate = $"«{DateTime.Now:dd}» {RussianMonthGenitive(DateTime.Now.Month)} {DateTime.Now.Year} г.";
                ReplaceParagraphText(body.Descendants<Paragraph>(),
                    text => text.Contains("г.") && text.Contains("«"),
                    _ => newDate);

                // ===== Таблица 1: Сдал / Принял =====
                var tables = body.Descendants<Table>().ToList();
                var infoTable = tables[0];
                var infoRows = infoTable.Elements<TableRow>().ToList();

                SetCellText(infoRows[0].Elements<TableCell>().ElementAt(1), handOverName);
                SetCellText(infoRows[0].Elements<TableCell>().ElementAt(2), $"ИИН: {handOverIIN}");
                SetCellText(infoRows[2].Elements<TableCell>().ElementAt(1), receiveName);
                SetCellText(infoRows[2].Elements<TableCell>().ElementAt(2), $"ИИН: {receiveIIN}");

                // ===== Таблица 2: список техники (первые 2 строки — заголовки) =====
                var devicesTable = tables[1];
                FillDevicesTable(devicesTable, TbLocation.Text.Trim());

                // ===== Таблица 3: блок подписей =====
                var sigTable = tables[2];
                var sigRows = sigTable.Elements<TableRow>().ToList();

                SetCellText(sigRows[0].Elements<TableCell>().ElementAt(1), handOverPosition);
                SetCellText(sigRows[0].Elements<TableCell>().ElementAt(5), ShortName(handOverName));
                SetCellText(sigRows[2].Elements<TableCell>().ElementAt(1), receivePosition);
                SetCellText(sigRows[2].Elements<TableCell>().ElementAt(5), ShortName(receiveName));

                doc.MainDocumentPart.Document.Save();
            }

            return outputPath;
        }

        // ===== Убираем недопустимые для имени файла символы, оставляя пусто -> "б-н" =====
        private static string SanitizeForFileName(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return "б-н";
            return string.Join("", value.Split(Path.GetInvalidFileNameChars()));
        }

        // ===== Загружаем docx-шаблон, вшитый в exe как Embedded Resource =====
        private static byte[] LoadTemplateBytes()
        {
            var assembly = Assembly.GetExecutingAssembly();
            string? resourceName = assembly.GetManifestResourceNames()
                .FirstOrDefault(n => n.EndsWith("Act_Template.docx", StringComparison.OrdinalIgnoreCase));

            if (resourceName == null)
                throw new FileNotFoundException(
                    "Шаблон акта Act_Template.docx не найден среди встроенных ресурсов проекта. " +
                    "Проверь, что файл добавлен с Build Action = Embedded Resource.");

            using var stream = assembly.GetManifestResourceStream(resourceName)!;
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }

        // ===== Заменяет весь текст параграфа (может состоять из нескольких <w:r>) одной строкой =====
        private static void ReplaceParagraphText(IEnumerable<Paragraph> paragraphs,
            Func<string, bool> match, Func<string, string> transform)
        {
            foreach (var p in paragraphs)
            {
                string text = string.Concat(p.Descendants<Text>().Select(t => t.Text));
                if (!match(text)) continue;

                var texts = p.Descendants<Text>().ToList();
                if (texts.Count == 0) continue;

                texts[0].Text = transform(text);
                texts[0].Space = SpaceProcessingModeValues.Preserve;
                for (int i = 1; i < texts.Count; i++)
                    texts[i].Text = "";
            }
        }

        // ===== Заменяет весь текст в ячейке таблицы одной строкой =====
        private static void SetCellText(TableCell cell, string newText)
        {
            var texts = cell.Descendants<Text>().ToList();
            if (texts.Count == 0) return;

            texts[0].Text = newText;
            texts[0].Space = SpaceProcessingModeValues.Preserve;
            for (int i = 1; i < texts.Count; i++)
                texts[i].Text = "";
        }

        // ===== Заполняет таблицу техники: 2 строки-заголовка остаются, дальше — по строке на устройство =====
        private void FillDevicesTable(Table devicesTable, string location)
        {
            var rows = devicesTable.Elements<TableRow>().ToList();
            const int headerRowCount = 2;

            // Строка-образец для клонирования — первая строка данных из шаблона
            TableRow templateRow = rows.Count > headerRowCount
                ? (TableRow)rows[headerRowCount].CloneNode(true)
                : throw new InvalidOperationException("В шаблоне не найдена строка данных для клонирования.");

            // Удаляем все существующие строки данных (оставляем только 2 заголовка)
            foreach (var row in rows.Skip(headerRowCount).ToList())
                row.Remove();

            var devicesToWrite = _devices.Count > 0
                ? _devices
                : new ObservableCollection<Device> { new Device() };

            int index = 1;
            foreach (var device in devicesToWrite)
            {
                var newRow = (TableRow)templateRow.CloneNode(true);
                var cells = newRow.Elements<TableCell>().ToList();

                SetCellText(cells[0], index.ToString());
                SetCellText(cells[1], device.Type);
                SetCellText(cells[2], device.Inv);
                SetCellText(cells[3], device.SN);
                SetCellText(cells[4], location);

                devicesTable.AppendChild(newRow);
                index++;
            }
        }

        // ===== Родительный падеж месяца для формата даты "01 июля 2026 г." =====
        private static string RussianMonthGenitive(int month) => month switch
        {
            1 => "января",
            2 => "февраля",
            3 => "марта",
            4 => "апреля",
            5 => "мая",
            6 => "июня",
            7 => "июля",
            8 => "августа",
            9 => "сентября",
            10 => "октября",
            11 => "ноября",
            12 => "декабря",
            _ => ""
        };

        // ===== "Аманжолов Мади" -> "Аманжолов М." для подписи =====
        private static string ShortName(string fullName)
        {
            var parts = fullName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            return parts.Length >= 2 ? $"{parts[0]} {parts[1][0]}." : fullName;
        }

        // ===== Анимация появления/исчезновения тоста (аналог showToast() из JS) =====
        private void ShowToast()
        {
            var appear = new Storyboard();

            var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200));
            Storyboard.SetTarget(fadeIn, Toast);
            Storyboard.SetTargetProperty(fadeIn, new PropertyPath(OpacityProperty));

            var slideIn = new DoubleAnimation(20, 0, TimeSpan.FromMilliseconds(200));
            Storyboard.SetTarget(slideIn, Toast);
            Storyboard.SetTargetProperty(slideIn, new PropertyPath("RenderTransform.Y"));

            appear.Children.Add(fadeIn);
            appear.Children.Add(slideIn);
            appear.Completed += (s, e) =>
            {
                // Через 3 секунды прячем тост обратно
                var hideTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(3)
                };
                hideTimer.Tick += (s2, e2) =>
                {
                    hideTimer.Stop();
                    var disappear = new Storyboard();

                    var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(300));
                    Storyboard.SetTarget(fadeOut, Toast);
                    Storyboard.SetTargetProperty(fadeOut, new PropertyPath(OpacityProperty));

                    var slideOut = new DoubleAnimation(0, 20, TimeSpan.FromMilliseconds(300));
                    Storyboard.SetTarget(slideOut, Toast);
                    Storyboard.SetTargetProperty(slideOut, new PropertyPath("RenderTransform.Y"));

                    disappear.Children.Add(fadeOut);
                    disappear.Children.Add(slideOut);
                    disappear.Begin();
                };
                hideTimer.Start();
            };

            appear.Begin();
        }
    }
}