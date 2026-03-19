using Project_bpi.Models;
using Project_bpi.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace Project_bpi
{
    public partial class ArchivePage : UserControl
    {
        private readonly string archiveFolderPath;
        private readonly Func<string, Task> downloadArchivedReportAsync;
        private readonly Func<string, Task> deleteArchivedReportAsync;
        private readonly List<ArchivedReportCard> allCards = new List<ArchivedReportCard>();
        private bool isLoaded;

        private sealed class ArchivedReportCard
        {
            public string Title { get; set; }
            public string DatabasePath { get; set; }
            public string ReportType { get; set; }
            public string ArchivedAtText { get; set; }
            public string SectionsText { get; set; }
            public string TablesText { get; set; }
        }

        public ArchivePage(
            string archiveFolderPath = null,
            Func<string, Task> downloadArchivedReportAsync = null,
            Func<string, Task> deleteArchivedReportAsync = null)
        {
            InitializeComponent();
            this.archiveFolderPath = string.IsNullOrWhiteSpace(archiveFolderPath)
                ? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ArchivedReports")
                : archiveFolderPath;
            this.downloadArchivedReportAsync = downloadArchivedReportAsync;
            this.deleteArchivedReportAsync = deleteArchivedReportAsync;

            Loaded += async (_, __) =>
            {
                if (isLoaded)
                {
                    return;
                }

                isLoaded = true;
                await ReloadAsync();
            };
        }

        public async Task ReloadAsync()
        {
            allCards.Clear();

            try
            {
                if (!Directory.Exists(archiveFolderPath))
                {
                    Directory.CreateDirectory(archiveFolderPath);
                }

                foreach (string databasePath in Directory
                    .GetFiles(archiveFolderPath, "*.db")
                    .OrderByDescending(File.GetLastWriteTimeUtc))
                {
                    var card = await TryLoadArchivedReportCardAsync(databasePath);
                    if (card != null)
                    {
                        allCards.Add(card);
                    }
                }
            }
            catch (Exception ex)
            {
                SummaryText.Text = $"Не удалось загрузить архив: {ex.Message}";
                EmptyStateText.Visibility = Visibility.Collapsed;
                ArchiveItemsControl.ItemsSource = null;
                return;
            }

            SummaryText.Text = allCards.Count == 0
                ? "Здесь отображаются сформированные итоговые отчеты."
                : $"В архиве отчетов: {allCards.Count}. Для каждой записи доступны скачивание в DOCX и удаление.";

            ApplyFilters();
        }

        private async Task<ArchivedReportCard> TryLoadArchivedReportCardAsync(string databasePath)
        {
            try
            {
                var database = new DataBase(databasePath);
                database.InitializeDatabase(false);

                var reports = await database.GetAllReports();
                var reportInfo = reports.FirstOrDefault();
                if (reportInfo == null)
                {
                    return null;
                }

                var report = await database.GetFullReport(reportInfo.Id);
                if (report == null)
                {
                    return null;
                }

                return new ArchivedReportCard
                {
                    Title = report.Title,
                    DatabasePath = databasePath,
                    ReportType = InferReportType(report.Title),
                    ArchivedAtText = $"Перенесен в архив: {File.GetLastWriteTime(databasePath):dd.MM.yyyy HH:mm}",
                    SectionsText = $"Разделов: {report.Sections?.Count ?? 0}",
                    TablesText = $"Таблиц: {CountTables(report)}"
                };
            }
            catch
            {
                return null;
            }
        }

        private string InferReportType(string title)
        {
            if (!string.IsNullOrWhiteSpace(title) && title.IndexOf("учеб", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Учебный отчет";
            }

            return "Отчет";
        }

        private int CountTables(Report report)
        {
            int total = 0;

            foreach (var section in report.Sections ?? Enumerable.Empty<Section>())
            {
                total += CountTables(section.SubSections);
            }

            return total;
        }

        private int CountTables(IEnumerable<SubSection> subSections)
        {
            int total = 0;

            foreach (var subsection in subSections ?? Enumerable.Empty<SubSection>())
            {
                total += subsection.Tables?.Count ?? 0;
                total += CountTables(subsection.SubSections);
            }

            return total;
        }

        private void Filter_Click(object sender, RoutedEventArgs e)
        {
            FilterPopup.IsOpen = true;
        }

        private void CancelFilter_Click(object sender, RoutedEventArgs e)
        {
            FilterPopup.IsOpen = false;
        }

        private void ApplyFilterFixed_Click(object sender, RoutedEventArgs e)
        {
            ApplyFilters();
            FilterPopup.IsOpen = false;
        }

        private void ApplyFilters()
        {
            string searchText = (FilterTextBox.Text ?? string.Empty).Trim();
            IEnumerable<ArchivedReportCard> filtered = allCards;

            if (RadioEducational.IsChecked == true)
            {
                filtered = filtered.Where(card => string.Equals(card.ReportType, "Учебный отчет", StringComparison.OrdinalIgnoreCase));
            }
            else if (RadioResearch.IsChecked == true)
            {
                filtered = filtered.Where(card => !string.Equals(card.ReportType, "Учебный отчет", StringComparison.OrdinalIgnoreCase));
            }

            if (!string.IsNullOrWhiteSpace(searchText))
            {
                filtered = filtered.Where(card =>
                    (card.Title?.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0 ||
                    (card.ReportType?.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) ?? -1) >= 0);
            }

            var result = filtered.ToList();
            ArchiveItemsControl.ItemsSource = result;
            EmptyStateText.Visibility = result.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void DownloadArchivedReportButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is ArchivedReportCard card))
            {
                return;
            }

            if (downloadArchivedReportAsync == null)
            {
                MessageBox.Show("Скачивание архивного отчета сейчас недоступно.", "Архив",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            button.IsEnabled = false;

            try
            {
                await downloadArchivedReportAsync(card.DatabasePath);
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        private async void DeleteArchivedReportButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is ArchivedReportCard card))
            {
                return;
            }

            if (MessageBox.Show(
                $"Удалить архивный отчет \"{card.Title}\"?",
                "Архив",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            button.IsEnabled = false;

            try
            {
                if (deleteArchivedReportAsync != null)
                {
                    await deleteArchivedReportAsync(card.DatabasePath);
                }
                else if (File.Exists(card.DatabasePath))
                {
                    File.Delete(card.DatabasePath);
                }

                await ReloadAsync();
            }
            finally
            {
                button.IsEnabled = true;
            }
        }
    }
}
