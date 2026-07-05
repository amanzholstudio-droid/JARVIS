using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace JARVIS
{
    // Читает файл матрицы (структура: 2 строки заголовков, дальше данные с 3-й строки).
    // Столбцы Место/Должность/Департамент/Управление заполнены только в первой строке
    // каждого блока (в Excel это выглядит как объединённые ячейки) — поэтому используем
    // "протягивание" последнего непустого значения для строк-продолжений.
    public static class MatrixLoader
    {
        public static List<MatrixEntry> Load(string path)
        {
            var result = new List<MatrixEntry>();

            using var workbook = new XLWorkbook(path);
            var sheet = workbook.Worksheets.First();

            string lastLocation = "", lastPosition = "", lastDept = "", lastDivision = "";

            int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;

            // Данные начинаются с 3-й строки: 1-я — объединённые групповые заголовки,
            // 2-я — заголовки колонок ("Должность", "ЦП" и т.д.)
            for (int row = 3; row <= lastRow; row++)
            {
                string position = sheet.Cell(row, 3).GetString().Trim();
                string location = sheet.Cell(row, 2).GetString().Trim();
                string dept = sheet.Cell(row, 4).GetString().Trim();
                string division = sheet.Cell(row, 5).GetString().Trim();

                // Новый блок начинается там, где явно указана должность
                if (!string.IsNullOrEmpty(position))
                {
                    lastLocation = location;
                    lastPosition = position;
                    lastDept = dept;
                    lastDivision = division;
                }

                string currentType = sheet.Cell(row, 6).GetString().Trim();

                // Пропускаем полностью пустые строки (случаются в конце файла)
                if (string.IsNullOrEmpty(lastPosition) || string.IsNullOrEmpty(currentType))
                    continue;

                result.Add(new MatrixEntry
                {
                    Location = lastLocation,
                    Position = lastPosition,
                    Department = lastDept,
                    Division = lastDivision,
                    CurrentType = currentType,
                    CurrentCpu = sheet.Cell(row, 7).GetString().Trim(),
                    CurrentRam = sheet.Cell(row, 8).GetString().Trim(),
                    CurrentSsd = sheet.Cell(row, 9).GetString().Trim(),
                    CurrentHdd = sheet.Cell(row, 10).GetString().Trim(),
                    CurrentGpu = sheet.Cell(row, 11).GetString().Trim(),
                    TargetType = sheet.Cell(row, 12).GetString().Trim(),
                    TargetCpu = sheet.Cell(row, 13).GetString().Trim(),
                    TargetRam = sheet.Cell(row, 14).GetString().Trim(),
                    TargetSsd = sheet.Cell(row, 15).GetString().Trim(),
                    TargetHdd = sheet.Cell(row, 16).GetString().Trim(),
                    TargetGpu = sheet.Cell(row, 17).GetString().Trim(),
                    Monitor = sheet.Cell(row, 18).GetString().Trim()
                });
            }

            return result;
        }
    }
}