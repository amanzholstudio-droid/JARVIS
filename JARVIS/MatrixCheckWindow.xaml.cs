using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Windows;
using System.Windows.Input;

namespace JARVIS
{
    public partial class MatrixCheckWindow : Window
    {
        private readonly List<MatrixEntry> _allEntries = new();
        private readonly ObservableCollection<MatrixEntry> _results = new();

        public MatrixCheckWindow(string? matrixPath)
        {
            InitializeComponent();
            ResultsList.ItemsSource = _results;

            if (string.IsNullOrWhiteSpace(matrixPath) || !File.Exists(matrixPath))
            {
                TbStatus.Text = "Матрица ещё не загружена. Загрузите файл через меню \"Загрузить матрицу\".";
                TbSearch.IsEnabled = false;
                return;
            }

            try
            {
                _allEntries.AddRange(MatrixLoader.Load(matrixPath));
                TbStatus.Text = $"Матрица загружена: {_allEntries.Count} строк. Введите должность и нажмите \"Найти\".";
            }
            catch (Exception ex)
            {
                TbStatus.Text = $"Не удалось прочитать файл матрицы: {ex.Message}";
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
                TbStatus.Text = "Введите должность для поиска.";
                return;
            }

            var matches = _allEntries
                .Where(entry => entry.Position.Contains(query, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (matches.Count == 0)
            {
                TbStatus.Text = $"«{query}» — НЕ ПОЛОЖЕНО. Должность не найдена в матрице.";
                return;
            }

            foreach (var entry in matches)
                _results.Add(entry);

            TbStatus.Text = $"«{query}» — ПОЛОЖЕНО. Найдено вариантов: {matches.Count}.";
        }
    }
}