using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace JARVIS
{
    public partial class JournalSearchWindow : Window
    {
        private readonly List<JournalRecord> _allRecords = new();
        private readonly ObservableCollection<JournalRecord> _results = new();

        public JournalSearchWindow(string? journalPath)
        {
            InitializeComponent();
            ResultsGrid.ItemsSource = _results;

            if (string.IsNullOrWhiteSpace(journalPath) || !File.Exists(journalPath))
            {
                TbStatus.Text = "Журнал ещё не создан — ни одной операции ещё не сохранено.";
                TbSearch.IsEnabled = false;
                return;
            }

            try
            {
                _allRecords.AddRange(JournalReader.Load(journalPath));
                TbStatus.Text = $"Журнал загружен: {_allRecords.Count} записей. Введи ИИН, инв. номер, серийный номер или ФИО.";
            }
            catch (System.Exception ex)
            {
                TbStatus.Text = $"Не удалось прочитать журнал: {ex.Message}";
                TbSearch.IsEnabled = false;
            }
        }

        private void TbSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) RunSearch();
        }

        private void Search_Click(object sender, RoutedEventArgs e) => RunSearch();

        private void RunSearch()
        {
            string query = TbSearch.Text.Trim();
            _results.Clear();

            if (string.IsNullOrWhiteSpace(query))
            {
                TbStatus.Text = "Введи запрос для поиска.";
                return;
            }

            var matches = _allRecords.Where(r =>
                Contains(r.EmployeeName, query) || Contains(r.EmployeeIIN, query) ||
                Contains(r.ToName, query) || Contains(r.ToIIN, query) ||
                Contains(r.Inv, query) || Contains(r.SN, query)
            ).ToList();

            foreach (var record in matches)
                _results.Add(record);

            TbStatus.Text = matches.Count == 0
                ? $"По запросу «{query}» ничего не найдено."
                : $"Найдено записей: {matches.Count}.";
        }

        private static bool Contains(string field, string query) =>
            field.Contains(query, StringComparison.OrdinalIgnoreCase);
    }
}