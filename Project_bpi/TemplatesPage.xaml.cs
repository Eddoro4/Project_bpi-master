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
    public partial class TemplatesPage : UserControl
    {
        private readonly string templatesFolderPath;
        private readonly Func<string, Task> unloadTemplateAsync;
        private readonly Func<string, Task> deleteTemplateAsync;
        private bool isLoaded;

        private sealed class SavedTemplateCard
        {
            public string Title { get; set; }
            public string DatabasePath { get; set; }
            public string SavedAtText { get; set; }
            public string SectionsText { get; set; }
            public string TablesText { get; set; }
        }

        public TemplatesPage(
            string templatesFolderPath = null,
            Func<string, Task> unloadTemplateAsync = null,
            Func<string, Task> deleteTemplateAsync = null)
        {
            InitializeComponent();
            this.templatesFolderPath = string.IsNullOrWhiteSpace(templatesFolderPath)
                ? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SavedTemplates")
                : templatesFolderPath;
            this.unloadTemplateAsync = unloadTemplateAsync;
            this.deleteTemplateAsync = deleteTemplateAsync;

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
            var cards = new List<SavedTemplateCard>();

            try
            {
                if (!Directory.Exists(templatesFolderPath))
                {
                    Directory.CreateDirectory(templatesFolderPath);
                }

                foreach (string databasePath in Directory
                    .GetFiles(templatesFolderPath, "*.db")
                    .OrderByDescending(File.GetLastWriteTimeUtc))
                {
                    var card = await TryLoadSavedTemplateCardAsync(databasePath);
                    if (card != null)
                    {
                        cards.Add(card);
                    }
                }
            }
            catch (Exception ex)
            {
                SummaryText.Text = $"Не удалось загрузить сохраненные шаблоны: {ex.Message}";
                EmptyStateText.Visibility = Visibility.Collapsed;
                SavedTemplatesItemsControl.ItemsSource = null;
                return;
            }

            SavedTemplatesItemsControl.ItemsSource = cards;
            EmptyStateText.Visibility = cards.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            SummaryText.Text = cards.Count == 0
                ? "Здесь хранятся неизменяемые снимки шаблонов. Их можно вернуть обратно в рабочее меню кнопкой «Выгрузить»."
                : $"Сохранено шаблонов: {cards.Count}. Снимки не редактируются и выгружаются в меню как отдельные рабочие копии.";
        }

        private async Task<SavedTemplateCard> TryLoadSavedTemplateCardAsync(string databasePath)
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

                int sectionsCount = report.Sections?.Count ?? 0;
                int tablesCount = CountTables(report);

                return new SavedTemplateCard
                {
                    Title = report.Title,
                    DatabasePath = databasePath,
                    SavedAtText = $"Сохранен: {File.GetLastWriteTime(databasePath):dd.MM.yyyy HH:mm}",
                    SectionsText = $"Разделов: {sectionsCount}",
                    TablesText = $"Таблиц: {tablesCount}"
                };
            }
            catch
            {
                return null;
            }
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

        private async void UnloadTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is SavedTemplateCard card))
            {
                return;
            }

            if (unloadTemplateAsync == null)
            {
                MessageBox.Show("Выгрузка шаблона сейчас недоступна.", "Шаблоны",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            button.IsEnabled = false;

            try
            {
                await unloadTemplateAsync(card.DatabasePath);
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        private async void DeleteTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is SavedTemplateCard card))
            {
                return;
            }

            if (MessageBox.Show(
                $"Удалить сохраненный шаблон \"{card.Title}\"?",
                "Шаблоны",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            button.IsEnabled = false;

            try
            {
                if (deleteTemplateAsync != null)
                {
                    await deleteTemplateAsync(card.DatabasePath);
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
