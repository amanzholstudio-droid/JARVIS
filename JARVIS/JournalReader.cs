using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace JARVIS
{
    // Читает journal.xlsx обратно в список записей.
    // Порядок колонок должен совпадать с JournalHeaders в MainWindow.xaml.cs:
    // Дата и время | Тип операции | Основание | № инцидента |
    // ФИО сотрудника | Должность | ИИН | Кому ФИО | Кому Должность | Кому ИИН |
    // Тип устройства | Инв. номер | Серийный номер
    public static class JournalReader
    {
        public static List<JournalRecord> Load(string path)
        {
            var result = new List<JournalRecord>();

            using var workbook = new XLWorkbook(path);
            var sheet = workbook.Worksheets.First();

            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;

            // Данные начинаются со 2-й строки (1-я — заголовки)
            for (int row = 2; row <= lastRow; row++)
            {
                string date = sheet.Cell(row, 1).GetString().Trim();
                if (string.IsNullOrEmpty(date)) continue; // пропускаем пустые строки

                result.Add(new JournalRecord
                {
                    DateTime = date,
                    OperationType = sheet.Cell(row, 2).GetString().Trim(),
                    Basis = sheet.Cell(row, 3).GetString().Trim(),
                    IncidentNo = sheet.Cell(row, 4).GetString().Trim(),
                    EmployeeName = sheet.Cell(row, 5).GetString().Trim(),
                    EmployeePosition = sheet.Cell(row, 6).GetString().Trim(),
                    EmployeeIIN = sheet.Cell(row, 7).GetString().Trim(),
                    ToName = sheet.Cell(row, 8).GetString().Trim(),
                    ToPosition = sheet.Cell(row, 9).GetString().Trim(),
                    ToIIN = sheet.Cell(row, 10).GetString().Trim(),
                    DeviceType = sheet.Cell(row, 11).GetString().Trim(),
                    Inv = sheet.Cell(row, 12).GetString().Trim(),
                    SN = sheet.Cell(row, 13).GetString().Trim()
                });
            }

            // Свежие операции сверху — удобнее смотреть последнюю историю
            result.Reverse();
            return result;
        }
    }
}