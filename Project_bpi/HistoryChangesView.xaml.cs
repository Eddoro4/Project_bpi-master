using Project_bpi.Models;
using Project_bpi.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Project_bpi
{
    public partial class HistoryChangesView : UserControl
    {
        private const string DefaultDatabaseFileName = "Kurs.db";
        private readonly string _databasePath;
        private List<HistoryEntryViewModel> _allEntries = new List<HistoryEntryViewModel>();
        private DateTime? _selectedStartDate;
        private DateTime? _selectedEndDate;
        private string _searchQuery = string.Empty;

        private sealed class HistoryEntryViewModel
        {
            public DateTime ChangedAtLocal { get; set; }
            public string ChangedAtDisplay { get; set; }
            public string Location { get; set; }
            public string Details { get; set; }
            public string ActionType { get; set; }
            public string ActionDisplay { get; set; }
            public Brush ActionBrush { get; set; }
            public Visibility DetailsVisibility => string.IsNullOrWhiteSpace(Details)
                ? Visibility.Collapsed
                : Visibility.Visible;
        }

        public HistoryChangesView(string databasePath = DefaultDatabaseFileName)
        {
            InitializeComponent();
            _databasePath = databasePath;
            Loaded += HistoryChangesView_Loaded;
        }

        private async void HistoryChangesView_Loaded(object sender, RoutedEventArgs e)
        {
            Loaded -= HistoryChangesView_Loaded;
            await LoadHistoryAsync();
        }

        private async Task LoadHistoryAsync()
        {
            var database = new DataBase(_databasePath);
            database.InitializeDatabase(false);

            var entries = await database.GetHistoryEntries(1000);
            _allEntries = entries.Select(MapHistoryEntry).ToList();
            ApplyFilters(showWarnings: false);
        }

        private HistoryEntryViewModel MapHistoryEntry(HistoryEntry entry)
        {
            string actionType = ResolveActionType(entry);
            return new HistoryEntryViewModel
            {
                ChangedAtLocal = entry.ChangedAtUtc.ToLocalTime(),
                ChangedAtDisplay = entry.ChangedAtUtc.ToLocalTime().ToString("dd.MM.yyyy HH:mm"),
                Location = NormalizeHistoryLocation(entry.Location),
                Details = entry.Details,
                ActionType = actionType,
                ActionDisplay = GetActionDisplay(actionType),
                ActionBrush = GetActionBrush(actionType)
            };
        }

        private static string ResolveActionType(HistoryEntry entry)
        {
            string actionType = (entry?.ActionType ?? string.Empty).Trim().ToLowerInvariant();
            string entityType = (entry?.EntityType ?? string.Empty).Trim().ToLowerInvariant();
            string details = (entry?.Details ?? string.Empty).Trim().ToLowerInvariant();

            if (actionType == "import" || entityType.Contains("import") || details.Contains("импорт"))
            {
                return "import";
            }

            if (actionType == "export" ||
                entityType.Contains("export") ||
                details.Contains("экспорт") ||
                details.Contains("выгруж"))
            {
                return "export";
            }

            return actionType;
        }

        private static string NormalizeHistoryLocation(string location)
        {
            if (string.IsNullOrWhiteSpace(location))
            {
                return string.Empty;
            }

            var normalizedParts = new List<string>();
            foreach (var part in location.Split(new[] { " / " }, System.StringSplitOptions.None))
            {
                string normalized = (part ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    continue;
                }

                if (normalizedParts.Count > 0 &&
                    string.Equals(normalizedParts[normalizedParts.Count - 1], normalized, System.StringComparison.Ordinal))
                {
                    continue;
                }

                normalizedParts.Add(normalized);
            }

            return string.Join(" / ", normalizedParts);
        }

        private static string GetActionDisplay(string actionType)
        {
            switch (actionType)
            {
                case "create":
                    return "Создано";
                case "import":
                    return "Импорт";
                case "export":
                    return "Экспорт";
                case "delete":
                    return "Удалено";
                default:
                    return "Отредактировано";
            }
        }

        private static Brush GetActionBrush(string actionType)
        {
            switch (actionType)
            {
                case "create":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1E7D3B"));
                case "import":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C27A00"));
                case "export":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0D8A72"));
                case "delete":
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C53A3A"));
                default:
                    return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF004AA1"));
            }
        }

        private void Filter_Click(object sender, RoutedEventArgs e)
        {
            FilterPopup.IsOpen = true;
        }

        private void ApplyFilter_Click(object sender, RoutedEventArgs e)
        {
            if (!ApplyFilters(showWarnings: true))
            {
                return;
            }

            FilterPopup.IsOpen = false;
        }

        private void ResetFilter_Click(object sender, RoutedEventArgs e)
        {
            CheckCreate.IsChecked = true;
            CheckEdit.IsChecked = true;
            CheckDelete.IsChecked = true;
            CheckImport.IsChecked = true;
            CheckExport.IsChecked = true;
            _selectedStartDate = null;
            _selectedEndDate = null;
            _searchQuery = string.Empty;
            SearchTextBox.Text = string.Empty;
            ApplyFilters(showWarnings: false);
            FilterPopup.IsOpen = false;
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            _searchQuery = (SearchTextBox.Text ?? string.Empty).Trim();
            ApplyFilters(showWarnings: false);
        }

        public void ApplyDateRangeFilter(DateTime startDate, DateTime endDate)
        {
            _selectedStartDate = startDate.Date;
            _selectedEndDate = endDate.Date;
            ApplyFilters(showWarnings: false);
        }

        private bool ApplyFilters(bool showWarnings)
        {
            var selectedActions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (CheckCreate.IsChecked == true)
            {
                selectedActions.Add("create");
            }

            if (CheckEdit.IsChecked == true)
            {
                selectedActions.Add("edit");
            }

            if (CheckDelete.IsChecked == true)
            {
                selectedActions.Add("delete");
            }

            if (CheckImport.IsChecked == true)
            {
                selectedActions.Add("import");
            }

            if (CheckExport.IsChecked == true)
            {
                selectedActions.Add("export");
            }

            if (selectedActions.Count == 0)
            {
                if (showWarnings)
                {
                    MessageBox.Show(
                        "Выберите хотя бы один тип действия.",
                        "История изменений",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }

                return false;
            }

            var filteredEntries = _allEntries
                .Where(entry => selectedActions.Contains(entry.ActionType))
                .Where(entry =>
                    (!_selectedStartDate.HasValue || entry.ChangedAtLocal.Date >= _selectedStartDate.Value) &&
                    (!_selectedEndDate.HasValue || entry.ChangedAtLocal.Date <= _selectedEndDate.Value))
                .Where(entry => string.IsNullOrWhiteSpace(_searchQuery) || MatchesSearch(entry, _searchQuery))
                .OrderByDescending(entry => entry.ChangedAtLocal)
                .ToList();

            HistoryItemsControl.ItemsSource = filteredEntries;
            EmptyStateText.Visibility = filteredEntries.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            SummaryText.Text = BuildSummaryText(filteredEntries.Count);

            return true;
        }

        private bool MatchesSearch(HistoryEntryViewModel entry, string query)
        {
            if (entry == null || string.IsNullOrWhiteSpace(query))
            {
                return true;
            }

            return ContainsIgnoreCase(entry.Location, query)
                || ContainsIgnoreCase(entry.Details, query)
                || ContainsIgnoreCase(entry.ActionDisplay, query)
                || ContainsIgnoreCase(entry.ChangedAtDisplay, query);
        }

        private bool ContainsIgnoreCase(string source, string query)
        {
            return !string.IsNullOrWhiteSpace(source)
                && source.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        private string BuildSummaryText(int count)
        {
            var fragments = new List<string>
            {
                $"Показано записей: {count}"
            };

            if (_selectedStartDate.HasValue && _selectedEndDate.HasValue)
            {
                fragments.Add($"период {_selectedStartDate.Value:dd.MM.yyyy} - {_selectedEndDate.Value:dd.MM.yyyy}");
            }

            if (!string.IsNullOrWhiteSpace(_searchQuery))
            {
                fragments.Add($"поиск: \"{_searchQuery}\"");
            }

            return string.Join(", ", fragments);
        }
    }
}
