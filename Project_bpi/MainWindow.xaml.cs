using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Globalization;
using Microsoft.Win32;
using Project_bpi.Models;
using Project_bpi.Services;

namespace Project_bpi
{
    public partial class MainWindow : Window
    {
        public DataBase DB = new DataBase();
        private const string SectionContentSubSectionTitle = "__section_content__";
        private const string NirReportTitle = "Отчет по НИР";
        private const string NirPublishingStaffReportTitle = "Издательская деятельность сотрудников кафедры и лаборатории";
        private const string NirPublishingTablePatternName = "nir_3_1_table";
        private const string NirPublishingStaffReportSuppressionKey = "builtin:nir_publishing_staff_report";
        private const string HistoryActionEdit = "edit";
        private const string HistoryActionImport = "import";
        private const string HistoryActionExport = "export";

        private sealed class TableCellSeed
        {
            public int Row { get; set; }
            public int Column { get; set; }
            public string Text { get; set; }
            public int ColSpan { get; set; } = 1;
            public int RowSpan { get; set; } = 1;
            public bool IsHeader { get; set; }
        }

        private sealed class TableSeed
        {
            public string Title { get; set; }
            public string PatternName { get; set; }
            public TableCellSeed[] Cells { get; set; }
        }

        private static readonly Lazy<IReadOnlyDictionary<string, HashSet<string>>> ReadOnlySeededBodyCellsByPattern =
            new Lazy<IReadOnlyDictionary<string, HashSet<string>>>(BuildReadOnlySeededBodyCellsByPattern);

        private sealed class NirSectionSeed
        {
            public int Number { get; set; }
            public string Title { get; set; }
            public TableSeed[] Tables { get; set; }
        }

        private readonly Dictionary<Border, Border> parents = new Dictionary<Border, Border>();
        private readonly Dictionary<Border, DynamicTemplateEntry> dynamicTemplates = new Dictionary<Border, DynamicTemplateEntry>();
        private readonly Dictionary<Border, Section> dynamicSections = new Dictionary<Border, Section>();
        private readonly Dictionary<Border, SubSection> dynamicSubSections = new Dictionary<Border, SubSection>();
        private readonly Dictionary<Border, StackPanel> dynamicMenus = new Dictionary<Border, StackPanel>();
        private readonly Dictionary<Border, Image> dynamicIndicators = new Dictionary<Border, Image>();
        private int _applicationZoomPercent = 100;
        private DateTime _currentCalendarDate = DateTime.Today;

        private sealed class DynamicTemplateEntry
        {
            public string DisplayTitle { get; set; }
            public string DatabasePath { get; set; }
            public int ReportId { get; set; }
            public Report Report { get; set; }
            public StackPanel Container { get; set; }
            public Border HeaderBorder { get; set; }
            public StackPanel MenuPanel { get; set; }
            public List<Border> RegisteredBorders { get; } = new List<Border>();
        }

        private sealed class TableEditorContext
        {
            public DynamicTemplateEntry Template { get; set; }
            public SubSection SubSection { get; set; }
            public Table Table { get; set; }
            public TextBox TitleTextBox { get; set; }
            public TableStructure Structure { get; set; }
            public ContentControl TableEditorHost { get; set; }
            public Func<TableEditorContext, UIElement> TableViewFactory { get; set; }
        }

        private const int MinimumApplicationZoomPercent = 50;
        private const int MaximumApplicationZoomPercent = 150;
        private const int ApplicationZoomStepPercent = 10;

        private sealed class Table7ExportRow
        {
            public string Number { get; set; }
            public string WorkName { get; set; }
            public string Performers { get; set; }
            public string PublicationType { get; set; }
            public string Volume { get; set; }
        }

        private sealed class NirPublicationExportRow
        {
            public string Number { get; set; }
            public string Share { get; set; }
            public string Authors { get; set; }
            public string PublicationName { get; set; }
            public string PublicationType { get; set; }
            public string EditionInfo { get; set; }
            public string PublicationPlace { get; set; }
        }

        private enum NirPublishingImportTemplate
        {
            PublicationsList,
            StudyPublishingPlan
        }

        public MainWindow()
        {
            DB.InitializeDatabase();
            InitializeComponent();
            ApplyApplicationZoom();
            InitializeDateRange();
            Loaded += async (sender, args) =>
            {
                await EnsureNirPublishingStaffReportInDatabaseAsync();
                await RemoveNirReportFromDatabaseAsync();
                await RemoveStudyReportFromDatabaseAsync();
                await LoadReportsAsync();
            };
        }

        private DynamicTemplateEntry AddTemplateToMenu(string templateTitle, Report report, string databasePath)
        {
            bool isDirectAccessTemplate = IsDirectAccessTemplateTitle(templateTitle);
            var templateContainer = new StackPanel();
            var templateMenu = new StackPanel { Visibility = Visibility.Collapsed };
            var templateHeader = CreateDynamicMenuBorder(
                templateTitle,
                "main",
                "MenuHeaderStyle",
                !isDirectAccessTemplate,
                out var templateIndicator);

            var entry = new DynamicTemplateEntry
            {
                DisplayTitle = templateTitle,
                DatabasePath = databasePath,
                ReportId = report.Id,
                Report = report
            };
            entry.Container = templateContainer;
            entry.HeaderBorder = templateHeader;
            entry.MenuPanel = templateMenu;

            dynamicTemplates[templateHeader] = entry;
            parents[templateHeader] = null;
            entry.RegisteredBorders.Add(templateHeader);

            if (!isDirectAccessTemplate)
            {
                dynamicMenus[templateHeader] = templateMenu;
                dynamicIndicators[templateHeader] = templateIndicator;
            }

            templateHeader.MouseLeftButtonDown += DynamicTemplateHeader_Click;

            templateContainer.Children.Add(templateHeader);

            if (!isDirectAccessTemplate)
            {
                templateContainer.Children.Add(templateMenu);
            }

            foreach (var section in report.Sections.OrderBy(s => s.Number))
            {
                section.Report = report;

                if (isDirectAccessTemplate)
                {
                    continue;
                }

                var sectionContainer = new StackPanel();
                var visibleSubSections = GetVisibleSubSections(section).ToList();
                bool hasChildren = visibleSubSections.Any();
                var sectionHeader = CreateDynamicMenuBorder(BuildSectionMenuTitle(section), "sub", "SubMenuStyle", hasChildren, out var sectionIndicator);

                dynamicSections[sectionHeader] = section;
                parents[sectionHeader] = templateHeader;
                sectionHeader.MouseLeftButtonDown += DynamicSectionHeader_Click;
                entry.RegisteredBorders.Add(sectionHeader);

                sectionContainer.Children.Add(sectionHeader);

                if (hasChildren)
                {
                    var sectionMenu = new StackPanel { Visibility = Visibility.Collapsed };
                    dynamicMenus[sectionHeader] = sectionMenu;
                    dynamicIndicators[sectionHeader] = sectionIndicator;
                    AddSubSectionsToMenu(sectionMenu, entry, section, visibleSubSections, sectionHeader, 0);

                    sectionContainer.Children.Add(sectionMenu);
                }

                templateMenu.Children.Add(sectionContainer);
            }

            InsertTemplateContainer(templateContainer, entry);
            return entry;
        }

        private void InsertTemplateContainer(StackPanel templateContainer, DynamicTemplateEntry entry)
        {
            if (templateContainer == null || entry == null)
            {
                return;
            }

            int newPriority = GetTemplateMenuPriority(entry.DisplayTitle);
            int insertIndex = DynamicTemplatesPanel.Children.Count;

            if (newPriority < int.MaxValue)
            {
                insertIndex = 0;

                foreach (var child in DynamicTemplatesPanel.Children.OfType<StackPanel>())
                {
                    var existingEntry = dynamicTemplates
                        .FirstOrDefault(pair => ReferenceEquals(pair.Value?.Container, child))
                        .Value;

                    if (existingEntry == null)
                    {
                        insertIndex++;
                        continue;
                    }

                    int existingPriority = GetTemplateMenuPriority(existingEntry.DisplayTitle);
                    if (existingPriority > newPriority)
                    {
                        break;
                    }

                    insertIndex++;
                }
            }

            DynamicTemplatesPanel.Children.Insert(insertIndex, templateContainer);
        }

        private int GetTemplateMenuPriority(string templateTitle)
        {
            if (string.Equals(templateTitle, NirPublishingStaffReportTitle, StringComparison.Ordinal))
            {
                return 0;
            }

            if (string.Equals(templateTitle, NirReportTitle, StringComparison.Ordinal))
            {
                return 1;
            }

            return int.MaxValue;
        }

        private void AddSubSectionsToMenu(
            Panel targetPanel,
            DynamicTemplateEntry entry,
            Section section,
            IEnumerable<SubSection> subSections,
            Border parentBorder,
            int depth)
        {
            foreach (var subsection in subSections.Where(item => !IsSectionContentSubSection(item)).OrderBy(s => s.Number))
            {
                subsection.Section = section;

                var subsectionContainer = new StackPanel();
                bool hasChildren = subsection.SubSections != null && subsection.SubSections.Any();
                var subsectionHeader = CreateDynamicMenuBorder(
                    BuildSubSectionTitle(section, subsection),
                    "sub2",
                    "SubMenuStyle1",
                    hasChildren,
                    out var subsectionIndicator);

                dynamicSubSections[subsectionHeader] = subsection;
                parents[subsectionHeader] = parentBorder;
                subsectionHeader.MouseLeftButtonDown += hasChildren
                    ? new MouseButtonEventHandler(DynamicSubSectionHeader_Click)
                    : DynamicLeaf_Click;
                entry.RegisteredBorders.Add(subsectionHeader);

                if (depth > 0)
                {
                    subsectionHeader.Padding = new Thickness(12 + depth * 12, 6, 12, 6);
                }

                subsectionContainer.Children.Add(subsectionHeader);

                if (hasChildren)
                {
                    var subsectionMenu = new StackPanel { Visibility = Visibility.Collapsed };
                    dynamicMenus[subsectionHeader] = subsectionMenu;
                    dynamicIndicators[subsectionHeader] = subsectionIndicator;

                    AddSubSectionsToMenu(subsectionMenu, entry, section, subsection.SubSections, subsectionHeader, depth + 1);
                    subsectionContainer.Children.Add(subsectionMenu);
                }

                targetPanel.Children.Add(subsectionContainer);
            }
        }

        private async Task LoadReportsAsync()
        {
            await LoadTemplatesFromDatabaseAsync(GetSharedTemplateDatabasePath());
        }

        private bool IsDirectAccessTemplate(DynamicTemplateEntry templateEntry)
        {
            return IsDirectAccessTemplateTitle(templateEntry?.DisplayTitle);
        }

        private bool IsDirectAccessTemplateTitle(string title)
        {
            return string.Equals(title, NirPublishingStaffReportTitle, StringComparison.Ordinal);
        }

        private Section GetDirectAccessSection(DynamicTemplateEntry templateEntry)
        {
            var section = templateEntry?.Report?.Sections?
                .OrderBy(item => item.Number)
                .FirstOrDefault();
            if (section != null)
            {
                section.Report = templateEntry.Report;
            }

            return section;
        }

        private async Task LoadTemplatesFromDatabaseAsync(string databasePath)
        {
            try
            {
                var database = new DataBase(databasePath);
                database.InitializeDatabase(false);

                var reports = await database.GetAllReports();
                foreach (var reportInfo in reports)
                {
                    if (!IsDirectAccessTemplateTitle(reportInfo.Title))
                    {
                        continue;
                    }

                    if (IsTemplateLoaded(databasePath, reportInfo.Id))
                    {
                        continue;
                    }

                    var report = await database.GetFullReport(reportInfo.Id);
                    if (report != null)
                    {
                        AddTemplateToMenu(report.Title, report, databasePath);
                    }
                }
            }
            catch
            {
                // Пропускаем поврежденную базу, чтобы остальные отчеты загрузились.
            }
        }

        private string GetSharedTemplateDatabasePath()
        {
            return DB.DatabasePath;
        }

        private async Task LogHistoryAsync(string actionType, string entityType, string location, string details = null)
        {
            if (string.IsNullOrWhiteSpace(location))
            {
                return;
            }

            try
            {
                var historyDatabase = new DataBase(GetSharedTemplateDatabasePath());
                historyDatabase.InitializeDatabase(false);
                await historyDatabase.AddHistoryEntry(new HistoryEntry
                {
                    ChangedAtUtc = DateTime.UtcNow,
                    ActionType = actionType,
                    EntityType = entityType ?? string.Empty,
                    Location = location.Trim(),
                    Details = details?.Trim()
                });
            }
            catch
            {
                // История не должна ломать основную работу приложения.
            }
        }

        private string BuildSectionHistoryTitle(Section section)
        {
            if (section == null)
            {
                return string.Empty;
            }

            string sectionContentTitle = GetSectionContentTitle(section);
            if (!string.IsNullOrWhiteSpace(sectionContentTitle))
            {
                return sectionContentTitle;
            }

            return BuildSectionMenuTitle(section);
        }

        private string BuildTemplateHistoryLocation(DynamicTemplateEntry templateEntry)
        {
            return templateEntry?.DisplayTitle?.Trim() ?? "Отчет";
        }

        private string BuildSectionHistoryLocation(DynamicTemplateEntry templateEntry, Section section)
        {
            return BuildHistoryLocation(
                BuildTemplateHistoryLocation(templateEntry),
                BuildSectionHistoryTitle(section));
        }

        private string BuildSubSectionHistoryLocation(DynamicTemplateEntry templateEntry, SubSection subsection)
        {
            if (subsection == null)
            {
                return BuildTemplateHistoryLocation(templateEntry);
            }

            var parts = new Stack<string>();
            var current = subsection;
            while (current != null)
            {
                if (!IsSectionContentSubSection(current) && !string.IsNullOrWhiteSpace(current.Title))
                {
                    parts.Push(current.Title.Trim());
                }

                current = current.ParentSubsection;
            }

            var locationParts = new List<string>
            {
                BuildTemplateHistoryLocation(templateEntry),
                BuildSectionHistoryTitle(subsection.Section)
            };

            locationParts.AddRange(parts);
            return BuildHistoryLocation(locationParts.ToArray());
        }

        private string BuildTableHistoryLocation(DynamicTemplateEntry templateEntry, SubSection subsection, Table table)
        {
            string tableTitle = string.IsNullOrWhiteSpace(table?.Title) ? "Таблица" : table.Title.Trim();
            if (subsection != null && IsSectionContentSubSection(subsection))
            {
                return BuildHistoryLocation(
                    BuildTemplateHistoryLocation(templateEntry),
                    BuildSectionHistoryTitle(subsection.Section),
                    tableTitle);
            }

            if (subsection == null)
            {
                return BuildHistoryLocation(
                    BuildTemplateHistoryLocation(templateEntry),
                    tableTitle);
            }

            var subsectionParts = new Stack<string>();
            var current = subsection;
            while (current != null)
            {
                if (!IsSectionContentSubSection(current) && !string.IsNullOrWhiteSpace(current.Title))
                {
                    subsectionParts.Push(current.Title.Trim());
                }

                current = current.ParentSubsection;
            }

            var locationParts = new List<string>
            {
                BuildTemplateHistoryLocation(templateEntry),
                BuildSectionHistoryTitle(subsection.Section)
            };

            locationParts.AddRange(subsectionParts);
            locationParts.Add(tableTitle);
            return BuildHistoryLocation(locationParts.ToArray());
        }

        private string BuildHistoryDetails(params string[] fragments)
        {
            return string.Join(", ", fragments.Where(fragment => !string.IsNullOrWhiteSpace(fragment)));
        }

        private string BuildHistoryLocation(params string[] parts)
        {
            var normalizedParts = new List<string>();

            foreach (var part in parts)
            {
                string normalized = part?.Trim();
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    continue;
                }

                if (normalizedParts.Count > 0 &&
                    string.Equals(normalizedParts[normalizedParts.Count - 1], normalized, StringComparison.Ordinal))
                {
                    continue;
                }

                normalizedParts.Add(normalized);
            }

            return string.Join(" / ", normalizedParts);
        }

        private bool IsTemplateLoaded(string databasePath, int reportId)
        {
            return dynamicTemplates.Values.Any(t =>
                string.Equals(t.DatabasePath, databasePath, StringComparison.OrdinalIgnoreCase) &&
                t.ReportId == reportId);
        }

        private string SanitizeFileName(string name)
        {
            string sanitized = Regex.Replace(name ?? "report", $"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()))}]", "_");
            return string.IsNullOrWhiteSpace(sanitized) ? "report" : sanitized.Trim();
        }

        private DynamicTemplateEntry GetCurrentTemplateEntry()
        {
            if (currentActive == null)
            {
                return null;
            }

            if (dynamicTemplates.TryGetValue(currentActive, out var templateEntry))
            {
                return templateEntry;
            }

            if (dynamicSections.TryGetValue(currentActive, out var section))
            {
                return FindTemplateEntryForSection(section);
            }

            if (dynamicSubSections.TryGetValue(currentActive, out var subsection))
            {
                return FindTemplateEntryForSubSection(subsection);
            }

            return null;
        }

        private TableEditorContext GetCurrentNirPublishingTableContext()
        {
            var templateEntry = GetCurrentTemplateEntry();
            if (templateEntry == null)
            {
                return null;
            }

            if (!string.Equals(templateEntry.DisplayTitle, NirPublishingStaffReportTitle, StringComparison.Ordinal))
            {
                return null;
            }

            var section = GetDirectAccessSection(templateEntry);
            var subsection = GetSectionContentSubSection(section);
            var table = subsection?.Tables?.FirstOrDefault(item =>
                string.Equals(item.PatternName, NirPublishingTablePatternName, StringComparison.Ordinal))
                ?? subsection?.Tables?.FirstOrDefault();

            if (table == null)
            {
                return null;
            }

            return new TableEditorContext
            {
                Template = templateEntry,
                SubSection = subsection,
                Table = table,
                Structure = CreateEditableTableStructure(table)
            };
        }

        private async void ExportByTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var context = GetCurrentNirPublishingTableContext();
            if (context == null)
            {
                MessageBox.Show("Выберите отчет \"Издательская деятельность сотрудников кафедры и лаборатории\".",
                    "Экспорт", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            await ExportNirPublishingTableByTemplate(context);
        }

        private void RemoveTemplateFromMenu(DynamicTemplateEntry entry)
        {
            foreach (var border in entry.RegisteredBorders)
            {
                dynamicTemplates.Remove(border);
                dynamicSections.Remove(border);
                dynamicSubSections.Remove(border);
                dynamicMenus.Remove(border);
                dynamicIndicators.Remove(border);
                parents.Remove(border);
            }

            entry.RegisteredBorders.Clear();

            if (entry.Container != null)
            {
                DynamicTemplatesPanel.Children.Remove(entry.Container);
            }
        }

        private async Task RefreshTemplateEntryAsync(
            DynamicTemplateEntry entry,
            int? sectionId = null,
            int? subsectionId = null)
        {
            var database = new DataBase(entry.DatabasePath);
            database.InitializeDatabase(false);

            Report report = await database.GetFullReport(entry.ReportId);
            if (report == null)
            {
                return;
            }

            RemoveTemplateFromMenu(entry);
            entry.Report = report;
            var refreshedEntry = AddTemplateToMenu(entry.DisplayTitle, report, entry.DatabasePath);

            Border selectedBorder = refreshedEntry.HeaderBorder;

            if (subsectionId.HasValue)
            {
                selectedBorder = dynamicSubSections.Keys.FirstOrDefault(border =>
                    dynamicSubSections.TryGetValue(border, out var subsection) &&
                    subsection.Id == subsectionId.Value &&
                    FindTemplateEntryForSubSection(subsection)?.ReportId == refreshedEntry.ReportId)
                    ?? selectedBorder;
            }
            else if (sectionId.HasValue)
            {
                selectedBorder = dynamicSections.Keys.FirstOrDefault(border =>
                    dynamicSections.TryGetValue(border, out var section) &&
                    section.Id == sectionId.Value &&
                    FindTemplateEntryForSection(section)?.ReportId == refreshedEntry.ReportId)
                    ?? selectedBorder;
            }

            if (selectedBorder != null)
            {
                ActivateMenuItem(selectedBorder);
            }
        }

        private async Task RemoveStudyReportFromDatabaseAsync()
        {
            var database = new DataBase(GetSharedTemplateDatabasePath());
            database.InitializeDatabase(false);

            var reports = await database.GetAllReports();
            foreach (var report in reports.Where(report =>
                string.Equals(report.Title, "Учебный отчет", StringComparison.Ordinal)).ToList())
            {
                await database.DeleteFilePattern(report.PattarnId);
            }
        }

        private async Task RemoveNirReportFromDatabaseAsync()
        {
            var database = new DataBase(GetSharedTemplateDatabasePath());
            database.InitializeDatabase(false);

            var reports = await database.GetAllReports();
            foreach (var report in reports.Where(report =>
                string.Equals(report.Title, NirReportTitle, StringComparison.Ordinal)).ToList())
            {
                await database.DeleteFilePattern(report.PattarnId);
            }
        }

        private Border CreateDynamicMenuBorder(string text, string tag, string textStyleKey, bool expandable, out Image indicator)
        {
            indicator = null;

            var border = new Border
            {
                Tag = tag,
                Style = (Style)FindResource("MenuItemStyle")
            };

            var title = new TextBlock
            {
                Text = text,
                Style = (Style)FindResource(textStyleKey),
                TextWrapping = TextWrapping.Wrap
            };

            if (!expandable)
            {
                border.Child = title;
                return border;
            }

            indicator = CreateDynamicMenuIndicator();

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var titlePanel = new StackPanel
            {
                Orientation = Orientation.Horizontal
            };
            titlePanel.Children.Add(title);
            titlePanel.Children.Add(indicator);

            grid.Children.Add(titlePanel);
            border.Child = grid;

            return border;
        }

        private void DynamicTemplateHeader_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border)
            {
                if (dynamicTemplates.TryGetValue(border, out var templateEntry) &&
                    IsDirectAccessTemplate(templateEntry))
                {
                    ActivateMenuItem(border);
                    return;
                }

                ToggleDynamicMenu(border);
                ActivateMenuItem(border, false);
            }
        }

        private void DynamicSectionHeader_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border)
            {
                if (dynamicMenus.ContainsKey(border))
                {
                    ToggleDynamicMenu(border);
                }

                ActivateMenuItem(border);
            }
        }

        private void DynamicLeaf_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border)
            {
                ActivateMenuItem(border);
            }
        }

        private void DynamicSubSectionHeader_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border)
            {
                if (dynamicMenus.ContainsKey(border))
                {
                    ToggleDynamicMenu(border);
                }

                ActivateMenuItem(border);
            }
        }

        private void ToggleDynamicMenu(Border border)
        {
            if (!dynamicMenus.TryGetValue(border, out var menu))
            {
                return;
            }

            bool isExpanded = menu.Visibility != Visibility.Visible;
            menu.Visibility = isExpanded ? Visibility.Visible : Visibility.Collapsed;

            if (dynamicIndicators.TryGetValue(border, out var indicator))
            {
                ApplyDynamicIndicatorState(indicator, isExpanded);
            }
        }

        private Image CreateDynamicMenuIndicator()
        {
            return new Image
            {
                Source = (ImageSource)FindResource("ArrowChevronBlueIcon"),
                Style = (Style)FindResource("ArrowIndicatorStyle")
            };
        }

        private void ApplyDynamicIndicatorState(Image indicator, bool isExpanded)
        {
            ApplyIndicatorVisual(indicator, isExpanded, false);
        }

        private void ApplyIndicatorVisual(Image indicator, bool isExpanded, bool isActive)
        {
            if (indicator == null)
            {
                return;
            }

            indicator.Source = (ImageSource)FindResource(isActive ? "ArrowChevronWhiteIcon" : "ArrowChevronBlueIcon");
            indicator.RenderTransform = isExpanded
                ? (Transform)new RotateTransform(-90)
                : Transform.Identity;
        }

        private string BuildSectionTitle(Section section)
        {
            if (section == null)
            {
                return string.Empty;
            }

            if (section.Number <= 0)
            {
                return string.IsNullOrWhiteSpace(section.Title)
                    ? string.Empty
                    : section.Title.Trim();
            }

            string defaultTitle = GetDefaultSectionTitle(section?.Number ?? 0);

            if (string.IsNullOrWhiteSpace(section.Title))
            {
                return defaultTitle;
            }

            if (string.Equals(section.Title.Trim(), defaultTitle, StringComparison.Ordinal))
            {
                return defaultTitle;
            }

            return $"Раздел {section.Number}. {section.Title}";
        }

        private string BuildSectionMenuTitle(Section section)
        {
            if (section != null && section.Number <= 0 && !string.IsNullOrWhiteSpace(section.Title))
            {
                return section.Title.Trim();
            }

            return GetDefaultSectionTitle(section?.Number ?? 0);
        }

        private string BuildSubSectionTitle(Section section, SubSection subsection)
        {
            if (string.IsNullOrWhiteSpace(subsection.Title))
            {
                return "Подраздел";
            }

            return subsection.Title.Trim();
        }

        private bool TryShowDynamicContent(Border menuItem)
        {
            if (dynamicTemplates.TryGetValue(menuItem, out var templateEntry))
            {
                var directSection = GetDirectAccessSection(templateEntry);
                MainContentControl.Content = directSection != null
                    ? CreateSectionPreviewContent(directSection)
                    : CreateDefaultContent();
                return true;
            }

            return false;
        }

        private UIElement CreateSectionPreviewContent(Section section)
        {
            var contentSubSection = GetSectionContentSubSection(section);
            var templateEntry = FindTemplateEntryForSection(section);
            string sectionText = GetSectionRawContent(section);
            string sectionTitle = GetSectionContentTitle(section);
            var stack = new StackPanel
            {
                Margin = new Thickness(20)
            };

            if (!string.IsNullOrWhiteSpace(sectionTitle))
            {
                stack.Children.Add(CreateSectionPreviewHeader(sectionTitle));
            }

            if (!string.IsNullOrWhiteSpace(sectionText))
            {
                stack.Children.Add(CreateSubSectionPreviewBody(sectionText));
            }

            if (contentSubSection?.Tables != null && contentSubSection.Tables.Any())
            {
                foreach (var table in contentSubSection.Tables.OrderBy(table => table.Id))
                {
                    stack.Children.Add(CreateSubSectionPreviewTableCard(templateEntry, contentSubSection, table));
                }
            }

            return new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = stack
            };
        }


        private string GetDefaultSectionTitle(int number)
        {
            return $"Раздел {number}";
        }

        private string GetSectionContentTitle(Section section)
        {
            if (section == null || string.IsNullOrWhiteSpace(section.Title))
            {
                return string.Empty;
            }

            if (string.Equals(section.Report?.Title, NirPublishingStaffReportTitle, StringComparison.Ordinal))
            {
                return NirPublishingStaffReportTitle;
            }

            string title = section.Title.Trim();
            if (section.Number <= 0)
            {
                return title;
            }

            return string.Equals(title, GetDefaultSectionTitle(section.Number), StringComparison.Ordinal)
                ? string.Empty
                : $"{section.Number} {title}";
        }

        private Border CreateSectionPreviewHeader(string sectionTitle)
        {
            var panel = new StackPanel
            {
                HorizontalAlignment = HorizontalAlignment.Center
            };

            panel.Children.Add(new TextBlock
            {
                Text = sectionTitle,
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                TextAlignment = TextAlignment.Center,
                TextWrapping = TextWrapping.Wrap
            });

            return new Border
            {
                CornerRadius = new CornerRadius(35),
                Margin = new Thickness(10),
                Padding = new Thickness(15),
                MinHeight = 80,
                Background = new LinearGradientBrush
                {
                    StartPoint = new Point(0, 0),
                    EndPoint = new Point(1, 1),
                    GradientStops = new GradientStopCollection
                    {
                        new GradientStop((Color)ColorConverter.ConvertFromString("#5394ba"), 0),
                        new GradientStop((Color)ColorConverter.ConvertFromString("#0167a4"), 1)
                    }
                },
                Child = panel
            };
        }

        private Border CreateSubSectionPreviewBody(string content)
        {
            return new Border
            {
                CornerRadius = new CornerRadius(35),
                Margin = new Thickness(10),
                Padding = new Thickness(15),
                Background = new LinearGradientBrush
                {
                    StartPoint = new Point(0, 0),
                    EndPoint = new Point(1, 1),
                    GradientStops = new GradientStopCollection
                    {
                        new GradientStop((Color)ColorConverter.ConvertFromString("#afc8d7"), 0)
                    }
                },
                Child = new ScrollViewer
                {
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    Content = new TextBox
                    {
                        Text = content,
                        FontSize = 12,
                        Foreground = Brushes.Black,
                        TextWrapping = TextWrapping.Wrap,
                        TextAlignment = TextAlignment.Justify,
                        AcceptsReturn = true,
                        AcceptsTab = false,
                        IsReadOnly = true,
                        Background = Brushes.Transparent,
                        BorderThickness = new Thickness(0),
                        Padding = new Thickness(0)
                    }
                }
            };
        }

        private Border CreateSubSectionPreviewTableCard(DynamicTemplateEntry templateEntry, SubSection subsection, Table table)
        {
            var structure = CreateEditableTableStructure(table);
            var previewHost = new ContentControl();
            var context = new TableEditorContext
            {
                Template = templateEntry,
                SubSection = subsection,
                Table = table,
                Structure = structure,
                TableEditorHost = previewHost,
                TableViewFactory = CreateFillableTableGrid
            };

            var content = new StackPanel();

            var titleRow = new Grid
            {
                Margin = new Thickness(0, 0, 0, 8)
            };

            var titleBlock = new TextBlock
            {
                Text = table.Title,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4")),
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center
            };
            titleRow.Children.Add(titleBlock);

            content.Children.Add(titleRow);

            previewHost.Content = CreateFillableTableGrid(context);
            content.Children.Add(previewHost);

            if (templateEntry != null && subsection != null)
            {
                content.Children.Add(new TextBlock
                {
                    Text = "Заполняйте строки прямо здесь и сохраняйте таблицу без перехода в режим редактирования.",
                    Margin = new Thickness(0, 8, 0, 8),
                    Foreground = Brushes.DimGray,
                    TextWrapping = TextWrapping.Wrap
                });

                var buttons = CreateButtonRow();

                var addRowButton = CreateSecondaryButton("Добавить строку");
                addRowButton.Tag = context;
                addRowButton.Click += AddTableRowButton_Click;
                buttons.Children.Add(addRowButton);

                if (CanImportNirPublishingTable(context))
                {
                    var importButton = CreateSecondaryButton("Импортировать таблицу");
                    importButton.Tag = context;
                    importButton.Click += ImportTableFromExcelButton_Click;
                    buttons.Children.Add(importButton);
                }

                var saveButton = CreateActionButton("Сохранить таблицу");
                saveButton.Tag = context;
                saveButton.Click += SavePreviewTableButton_Click;
                buttons.Children.Add(saveButton);

                var clearButton = CreateSecondaryButton("Очистить таблицу");
                clearButton.Tag = context;
                clearButton.Click += ClearTableContentButton_Click;
                buttons.Children.Add(clearButton);

                content.Children.Add(buttons);
            }

            return new Border
            {
                CornerRadius = new CornerRadius(24),
                Margin = new Thickness(10, 2, 10, 10),
                Padding = new Thickness(18, 14, 18, 14),
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#d6e2ea")),
                BorderThickness = new Thickness(1),
                Child = content
            };
        }

        private StackPanel CreateContentStack(string title, string subtitle)
        {
            var stack = new StackPanel
            {
                Margin = new Thickness(24)
            };

            if (!string.IsNullOrWhiteSpace(title))
            {
                stack.Children.Add(new TextBlock
                {
                    Text = title,
                    FontSize = 22,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4")),
                    TextWrapping = TextWrapping.Wrap
                });
            }

            stack.Children.Add(new TextBlock
            {
                Text = subtitle,
                Margin = new Thickness(0, string.IsNullOrWhiteSpace(title) ? 0 : 8, 0, 18),
                Foreground = Brushes.DimGray,
                TextWrapping = TextWrapping.Wrap
            });

            return stack;
        }

        private UIElement CreateDefaultContent()
        {
            return new TextBlock
            {
                Text = "Основная область контента",
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = Brushes.Gray,
                FontSize = 16
            };
        }

        private DynamicTemplateEntry FindTemplateEntryForSection(Section section)
        {
            return dynamicTemplates.Values.FirstOrDefault(entry =>
                entry.Report.Sections.Contains(section));
        }

        private DynamicTemplateEntry FindTemplateEntryForSubSection(SubSection subsection)
        {
            return dynamicTemplates.Values.FirstOrDefault(entry =>
                entry.Report.Sections.Any(section => ContainsSubSection(section.SubSections, subsection)));
        }

        private bool IsSectionContentSubSection(SubSection subsection)
        {
            return subsection != null
                && !subsection.ParentSubsectionId.HasValue
                && string.Equals(subsection.Title, SectionContentSubSectionTitle, System.StringComparison.Ordinal);
        }

        private IEnumerable<SubSection> GetVisibleSubSections(Section section)
        {
            return section?.SubSections?.Where(item => !IsSectionContentSubSection(item))
                ?? Enumerable.Empty<SubSection>();
        }

        private SubSection GetSectionContentSubSection(Section section)
        {
            return section?.SubSections?.FirstOrDefault(IsSectionContentSubSection);
        }

        private bool ContainsSubSection(IEnumerable<SubSection> subSections, SubSection target)
        {
            if (subSections == null || target == null)
            {
                return false;
            }

            foreach (var subsection in subSections)
            {
                if (ReferenceEquals(subsection, target) || subsection.Id == target.Id)
                {
                    return true;
                }

                if (subsection.SubSections != null && ContainsSubSection(subsection.SubSections, target))
                {
                    return true;
                }
            }

            return false;
        }

        private IEnumerable<string> BuildSubSectionOutlineLines(Section section, IEnumerable<SubSection> subSections, int depth = 0)
        {
            if (subSections == null)
            {
                yield break;
            }

            foreach (var subsection in subSections.Where(item => !IsSectionContentSubSection(item)).OrderBy(item => item.Number))
            {
                string indent = new string(' ', depth * 4);
                yield return indent + BuildSubSectionTitle(section, subsection);

                if (subsection.SubSections == null)
                {
                    continue;
                }

                foreach (var line in BuildSubSectionOutlineLines(section, subsection.SubSections, depth + 1))
                {
                    yield return line;
                }
            }
        }

        private TextBox CreateEditorTextBox(string text, bool multiline)
        {
            return new TextBox
            {
                Text = text ?? string.Empty,
                AcceptsReturn = multiline,
                TextWrapping = TextWrapping.Wrap,
                MinHeight = multiline ? 140 : 36,
                Padding = new Thickness(10, 8, 10, 8),
                VerticalContentAlignment = VerticalAlignment.Top
            };
        }

        private StackPanel CreateButtonRow()
        {
            return new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 12)
            };
        }

        private Button CreateActionButton(string text)
        {
            return new Button
            {
                Content = text,
                Margin = new Thickness(0, 0, 10, 12),
                Padding = new Thickness(14, 8, 14, 8),
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4")),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
        }

        private Button CreateSecondaryButton(string text)
        {
            return new Button
            {
                Content = text,
                Margin = new Thickness(0, 0, 10, 12),
                Padding = new Thickness(14, 8, 14, 8),
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#e9ecef")),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#212529")),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ced4da")),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand
            };
        }

        private Button CreateDangerButton(string text)
        {
            return new Button
            {
                Content = text,
                Margin = new Thickness(0, 0, 10, 12),
                Padding = new Thickness(14, 8, 14, 8),
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#b02a37")),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
        }

        private TableStructure CreateEditableTableStructure(Table table)
        {
            var structure = ExtractTableStructure(table);
            EnsureEditableTableStructure(structure);
            return structure;
        }

        private void RefreshTableEditor(TableEditorContext context)
        {
            if (context?.TableEditorHost == null)
            {
                return;
            }

            EnsureEditableTableStructure(context.Structure);
            context.TableEditorHost.Content = (context.TableViewFactory ?? CreateFillableTableGrid)(context);
        }

        private void EnsureEditableTableStructure(TableStructure structure)
        {
            if (structure == null)
            {
                return;
            }

            NormalizeTableStructure(structure);

            if (structure.HeaderRowCount <= 0)
            {
                structure.HeaderRowCount = 1;
            }

            if (structure.ColumnCount <= 0)
            {
                structure.ColumnCount = 1;
            }

            EnsureEditableCells(structure.HeaderCells, structure.HeaderRowCount, structure.ColumnCount, true);
            EnsureEditableCells(structure.BodyCells, structure.BodyRowCount, structure.ColumnCount, false);
            NormalizeTableStructure(structure);
        }

        private void EnsureEditableCells(List<TableCellDefinition> cells, int rowCount, int columnCount, bool isHeader)
        {
            if (rowCount <= 0 || columnCount <= 0)
            {
                return;
            }

            var occupied = new bool[rowCount + 1, columnCount + 1];
            foreach (var cell in cells)
            {
                for (int row = cell.Row; row < cell.Row + cell.RowSpan && row <= rowCount; row++)
                {
                    for (int column = cell.Column; column < cell.Column + cell.ColSpan && column <= columnCount; column++)
                    {
                        occupied[row, column] = true;
                    }
                }
            }

            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    if (occupied[row, column])
                    {
                        continue;
                    }

                    cells.Add(new TableCellDefinition
                    {
                        Text = string.Empty,
                        Row = row,
                        Column = column,
                        ColSpan = 1,
                        RowSpan = 1,
                        IsHeader = isHeader
                    });
                }
            }
        }

        private UIElement CreateFillableTableGrid(TableEditorContext context)
        {
            EnsureEditableTableStructure(context.Structure);

            var structure = context.Structure;
            int columnCount = structure.ColumnCount;
            var visibleBodyRows = GetVisibleBodyRows(context);

            if (columnCount == 0)
            {
                return new Border
                {
                    Background = Brushes.White,
                    BorderBrush = Brushes.Black,
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(12),
                    Child = new TextBlock
                    {
                        Text = "Таблица пока пустая.",
                        Foreground = Brushes.Gray
                    }
                };
            }

            var panel = new StackPanel
            {
                Background = Brushes.White
            };

            if (structure.HeaderRowCount > 0)
            {
                var headerGrid = new Grid();
                ConfigurePreviewTableColumns(headerGrid, columnCount);

                for (int rowIndex = 0; rowIndex < structure.HeaderRowCount; rowIndex++)
                {
                    headerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                }

                AddHeaderCells(headerGrid, structure, columnCount);
                panel.Children.Add(headerGrid);
            }

            if (structure.BodyRowCount > 0)
            {
                var bodyGrid = new Grid();
                ConfigurePreviewTableColumns(bodyGrid, columnCount);

                for (int rowIndex = 0; rowIndex < visibleBodyRows.Count; rowIndex++)
                {
                    bodyGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                }

                if (visibleBodyRows.Count > 0)
                {
                    AddFillablePreviewCellsToGrid(bodyGrid, context, structure.BodyCells, visibleBodyRows, columnCount);
                    panel.Children.Add(bodyGrid);
                }
                else
                {
                    panel.Children.Add(CreateEmptyTableMessageBorder("Строки таблицы отсутствуют."));
                }
            }
            else
            {
                panel.Children.Add(CreateEmptyTableMessageBorder("Строки данных отсутствуют. Добавьте строку, чтобы заполнить таблицу."));
            }

            var tableScrollViewer = new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = panel
            };
            AttachMouseWheelScrolling(tableScrollViewer);

            return new Border
            {
                Background = Brushes.White,
                BorderBrush = Brushes.Black,
                BorderThickness = new Thickness(1),
                ClipToBounds = true,
                Child = tableScrollViewer
            };
        }

        private Border CreateEmptyTableMessageBorder(string message)
        {
            return new Border
            {
                BorderBrush = Brushes.Black,
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(10),
                Child = new TextBlock
                {
                    Text = message,
                    Foreground = Brushes.Gray,
                    TextWrapping = TextWrapping.Wrap
                }
            };
        }

        private void AttachMouseWheelScrolling(UIElement element)
        {
            if (element == null)
            {
                return;
            }

            element.PreviewMouseWheel -= TableScrollViewer_PreviewMouseWheel;
            element.PreviewMouseWheel += TableScrollViewer_PreviewMouseWheel;
        }

        private void TableScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (!(sender is DependencyObject dependencyObject) || e == null || e.Delta == 0)
            {
                return;
            }

            double linesPerWheelStep = SystemParameters.WheelScrollLines > 0
                ? SystemParameters.WheelScrollLines
                : 3d;
            double offsetStep = linesPerWheelStep * 16d;
            bool isShiftPressed = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;

            if (isShiftPressed)
            {
                ScrollViewer horizontalScrollViewer = FindScrollableHorizontalParent(dependencyObject, e.Delta);
                if (horizontalScrollViewer == null)
                {
                    return;
                }

                double targetHorizontalOffset = e.Delta > 0
                    ? horizontalScrollViewer.HorizontalOffset - offsetStep
                    : horizontalScrollViewer.HorizontalOffset + offsetStep;

                horizontalScrollViewer.ScrollToHorizontalOffset(
                    Math.Max(0d, Math.Min(horizontalScrollViewer.ScrollableWidth, targetHorizontalOffset)));
                e.Handled = true;
                return;
            }

            ScrollViewer targetScrollViewer = FindScrollableVerticalParent(dependencyObject, e.Delta);
            if (targetScrollViewer == null)
            {
                return;
            }

            double targetOffset = e.Delta > 0
                ? targetScrollViewer.VerticalOffset - offsetStep
                : targetScrollViewer.VerticalOffset + offsetStep;

            targetScrollViewer.ScrollToVerticalOffset(
                Math.Max(0d, Math.Min(targetScrollViewer.ScrollableHeight, targetOffset)));
            e.Handled = true;
        }

        private ScrollViewer FindScrollableVerticalParent(DependencyObject start, int delta)
        {
            ScrollViewer fallbackScrollViewer = null;

            for (DependencyObject current = start; current != null; current = GetVisualOrLogicalParent(current))
            {
                if (!(current is ScrollViewer scrollViewer) || scrollViewer.ScrollableHeight <= 0)
                {
                    continue;
                }

                if (fallbackScrollViewer == null)
                {
                    fallbackScrollViewer = scrollViewer;
                }

                bool canScrollUp = scrollViewer.VerticalOffset > 0;
                bool canScrollDown = scrollViewer.VerticalOffset < scrollViewer.ScrollableHeight;
                if ((delta > 0 && canScrollUp) || (delta < 0 && canScrollDown))
                {
                    return scrollViewer;
                }
            }

            return fallbackScrollViewer;
        }

        private ScrollViewer FindScrollableHorizontalParent(DependencyObject start, int delta)
        {
            ScrollViewer fallbackScrollViewer = null;

            for (DependencyObject current = start; current != null; current = GetVisualOrLogicalParent(current))
            {
                if (!(current is ScrollViewer scrollViewer) || scrollViewer.ScrollableWidth <= 0)
                {
                    continue;
                }

                if (fallbackScrollViewer == null)
                {
                    fallbackScrollViewer = scrollViewer;
                }

                bool canScrollLeft = scrollViewer.HorizontalOffset > 0;
                bool canScrollRight = scrollViewer.HorizontalOffset < scrollViewer.ScrollableWidth;
                if ((delta > 0 && canScrollLeft) || (delta < 0 && canScrollRight))
                {
                    return scrollViewer;
                }
            }

            return fallbackScrollViewer;
        }

        private DependencyObject GetVisualOrLogicalParent(DependencyObject element)
        {
            if (element == null)
            {
                return null;
            }

            if (element is Visual || element is System.Windows.Media.Media3D.Visual3D)
            {
                DependencyObject visualParent = VisualTreeHelper.GetParent(element);
                if (visualParent != null)
                {
                    return visualParent;
                }
            }

            if (element is FrameworkElement frameworkElement)
            {
                return frameworkElement.Parent;
            }

            if (element is FrameworkContentElement frameworkContentElement)
            {
                return frameworkContentElement.Parent;
            }

            return null;
        }

        private void ApplyApplicationZoom()
        {
            double zoomScale = Math.Max(
                MinimumApplicationZoomPercent,
                Math.Min(MaximumApplicationZoomPercent, _applicationZoomPercent)) / 100d;

            if (ApplicationScaleTransform != null)
            {
                ApplicationScaleTransform.ScaleX = zoomScale;
                ApplicationScaleTransform.ScaleY = zoomScale;
            }

            if (AppZoomPercentText != null)
            {
                AppZoomPercentText.Text = $"{_applicationZoomPercent}%";
            }

            if (DecreaseAppZoomButton != null)
            {
                DecreaseAppZoomButton.IsEnabled = _applicationZoomPercent > MinimumApplicationZoomPercent;
            }

            if (IncreaseAppZoomButton != null)
            {
                IncreaseAppZoomButton.IsEnabled = _applicationZoomPercent < MaximumApplicationZoomPercent;
            }

            if (ResetAppZoomButton != null)
            {
                ResetAppZoomButton.IsEnabled = _applicationZoomPercent != 100;
            }
        }

        private void ChangeApplicationZoom(int delta)
        {
            int newZoomPercent = Math.Max(
                MinimumApplicationZoomPercent,
                Math.Min(MaximumApplicationZoomPercent, _applicationZoomPercent + delta));

            if (newZoomPercent == _applicationZoomPercent)
            {
                return;
            }

            _applicationZoomPercent = newZoomPercent;
            ApplyApplicationZoom();
        }

        private void DecreaseAppZoomButton_Click(object sender, RoutedEventArgs e)
        {
            ChangeApplicationZoom(-ApplicationZoomStepPercent);
        }

        private void IncreaseAppZoomButton_Click(object sender, RoutedEventArgs e)
        {
            ChangeApplicationZoom(ApplicationZoomStepPercent);
        }

        private void ResetAppZoomButton_Click(object sender, RoutedEventArgs e)
        {
            if (_applicationZoomPercent == 100)
            {
                return;
            }

            _applicationZoomPercent = 100;
            ApplyApplicationZoom();
        }

        private IReadOnlyList<int> GetVisibleBodyRows(TableEditorContext context)
        {
            if (context?.Structure == null || context.Structure.BodyRowCount <= 0)
            {
                return Array.Empty<int>();
            }

            return Enumerable.Range(1, context.Structure.BodyRowCount).ToList();
        }

        private string GetBodyCellText(TableStructure structure, int bodyRow, int columnIndex)
        {
            var cell = structure?.BodyCells.FirstOrDefault(item =>
                bodyRow >= item.Row &&
                bodyRow < item.Row + item.RowSpan &&
                columnIndex >= item.Column &&
                columnIndex < item.Column + item.ColSpan);

            return cell?.Text?.Trim() ?? string.Empty;
        }

        private string GetSectionRawContent(Section section)
        {
            var contentSubSection = GetSectionContentSubSection(section);
            return contentSubSection?.Texts != null && contentSubSection.Texts.Any()
                ? contentSubSection.Texts.First().Content
                : string.Empty;
        }

        private void AddFillablePreviewCellsToGrid(
            Grid grid,
            TableEditorContext context,
            IReadOnlyCollection<TableCellDefinition> cells,
            IReadOnlyList<int> visibleRows,
            int columnCount)
        {
            var readOnlyCells = GetReadOnlySeededBodyCells(context?.Table);
            var rowMap = visibleRows
                .Select((originalRow, index) => new { originalRow, displayRow = index + 1 })
                .ToDictionary(item => item.originalRow, item => item.displayRow);
            var occupied = new bool[visibleRows.Count + 1, columnCount + 1];

            foreach (var cell in cells.OrderBy(item => item.Row).ThenBy(item => item.Column))
            {
                if (!rowMap.TryGetValue(cell.Row, out int displayRow))
                {
                    continue;
                }

                bool shouldRenderAsHeader = cell.IsHeader;
                UIElement element = shouldRenderAsHeader || readOnlyCells.Contains(GetTableCellSeedKey(cell))
                    ? (UIElement)CreateTableCellBorder(cell.Text, cell.IsHeader)
                    : CreateEditableTableCellTextBox(cell);

                Grid.SetRow(element, displayRow - 1);
                Grid.SetColumn(element, cell.Column - 1);
                Grid.SetColumnSpan(element, cell.ColSpan);
                Grid.SetRowSpan(element, 1);
                grid.Children.Add(element);

                for (int column = cell.Column; column < cell.Column + cell.ColSpan; column++)
                {
                    occupied[displayRow, column] = true;
                }
            }

            for (int row = 1; row <= visibleRows.Count; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    if (occupied[row, column])
                    {
                        continue;
                    }

                    var emptyBorder = CreateTableCellBorder(string.Empty, false);
                    Grid.SetRow(emptyBorder, row - 1);
                    Grid.SetColumn(emptyBorder, column - 1);
                    grid.Children.Add(emptyBorder);
                }
            }
        }

        private static IReadOnlyDictionary<string, HashSet<string>> BuildReadOnlySeededBodyCellsByPattern()
        {
            var result = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);

            foreach (var seed in GetAllKnownTableSeeds())
            {
                if (seed?.Cells == null || string.IsNullOrWhiteSpace(seed.PatternName))
                {
                    continue;
                }

                if (!result.TryGetValue(seed.PatternName, out var cellKeys))
                {
                    cellKeys = new HashSet<string>(StringComparer.Ordinal);
                    result[seed.PatternName] = cellKeys;
                }

                foreach (var cell in seed.Cells.Where(item => item != null && !item.IsHeader && !string.IsNullOrWhiteSpace(item.Text)))
                {
                    cellKeys.Add(GetTableCellSeedKey(cell));
                }
            }

            return result;
        }

        private static IEnumerable<TableSeed> GetAllKnownTableSeeds()
        {
            foreach (var table in CreateNirPublishingStaffReportSection().Tables ?? Array.Empty<TableSeed>())
            {
                yield return table;
            }
        }

        private static HashSet<string> GetReadOnlySeededBodyCells(Table table)
        {
            if (table == null || string.IsNullOrWhiteSpace(table.PatternName))
            {
                return EmptyReadOnlySeededBodyCells;
            }

            return ReadOnlySeededBodyCellsByPattern.Value.TryGetValue(table.PatternName, out var cells)
                ? cells
                : EmptyReadOnlySeededBodyCells;
        }

        private static readonly HashSet<string> EmptyReadOnlySeededBodyCells =
            new HashSet<string>(StringComparer.Ordinal);

        private static string GetTableCellSeedKey(TableCellDefinition cell)
        {
            return GetTableCellSeedKey(cell.Row, cell.Column, cell.ColSpan, cell.RowSpan, cell.Text);
        }

        private static string GetTableCellSeedKey(TableCellSeed cell)
        {
            return GetTableCellSeedKey(cell.Row, cell.Column, cell.ColSpan, cell.RowSpan, cell.Text);
        }

        private static string GetTableCellSeedKey(int row, int column, int colSpan, int rowSpan, string text)
        {
            return string.Concat(
                row.ToString(CultureInfo.InvariantCulture), "|",
                column.ToString(CultureInfo.InvariantCulture), "|",
                colSpan.ToString(CultureInfo.InvariantCulture), "|",
                rowSpan.ToString(CultureInfo.InvariantCulture), "|",
                text ?? string.Empty);
        }

        private TextBox CreateEditableTableCellTextBox(
            TableCellDefinition cell)
        {
            var textBox = new TextBox
            {
                Text = cell.Text ?? string.Empty,
                MinWidth = 0,
                MinHeight = cell.IsHeader ? 56 : 48,
                Padding = new Thickness(10, 8, 10, 8),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#b6c2cf")),
                Background = cell.IsHeader
                    ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eef6ff"))
                    : Brushes.White,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1f2933")),
                FontWeight = cell.IsHeader ? FontWeights.SemiBold : FontWeights.Normal,
                TextWrapping = TextWrapping.Wrap,
                AcceptsReturn = true,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalContentAlignment = cell.IsHeader ? HorizontalAlignment.Center : HorizontalAlignment.Left,
                TextAlignment = cell.IsHeader ? TextAlignment.Center : TextAlignment.Left,
                ClipToBounds = true
            };

            textBox.TextChanged += (sender, args) =>
            {
                cell.Text = textBox.Text;
            };
            textBox.PreviewKeyDown += EditableTableCellTextBox_PreviewKeyDown;

            return textBox;
        }

        private void EditableTableCellTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!(sender is TextBox currentTextBox))
            {
                return;
            }

            if (e.Key != Key.Left &&
                e.Key != Key.Right &&
                e.Key != Key.Up &&
                e.Key != Key.Down)
            {
                return;
            }

            if (!ShouldMoveToAdjacentTableCell(currentTextBox, e.Key))
            {
                return;
            }

            var targetTextBox = FindAdjacentTableCellTextBox(currentTextBox, e.Key);
            if (targetTextBox == null)
            {
                return;
            }

            e.Handled = true;
            MoveCaretToAdjacentTableCell(currentTextBox, targetTextBox, e.Key);
        }

        private TextBox FindAdjacentTableCellTextBox(TextBox currentTextBox, Key directionKey)
        {
            if (!(currentTextBox.Parent is Grid parentGrid))
            {
                return null;
            }

            int currentRow = Grid.GetRow(currentTextBox);
            int currentColumn = Grid.GetColumn(currentTextBox);
            int rowSpan = Math.Max(1, Grid.GetRowSpan(currentTextBox));
            int columnSpan = Math.Max(1, Grid.GetColumnSpan(currentTextBox));

            int rowStep = 0;
            int columnStep = 0;
            int nextRow = currentRow;
            int nextColumn = currentColumn;

            switch (directionKey)
            {
                case Key.Left:
                    columnStep = -1;
                    nextColumn = currentColumn - 1;
                    break;
                case Key.Right:
                    columnStep = 1;
                    nextColumn = currentColumn + columnSpan;
                    break;
                case Key.Up:
                    rowStep = -1;
                    nextRow = currentRow - 1;
                    break;
                case Key.Down:
                    rowStep = 1;
                    nextRow = currentRow + rowSpan;
                    break;
                default:
                    return null;
            }

            int maxRow = Math.Max(0, parentGrid.RowDefinitions.Count - 1);
            int maxColumn = Math.Max(0, parentGrid.ColumnDefinitions.Count - 1);

            while (nextRow >= 0 && nextRow <= maxRow && nextColumn >= 0 && nextColumn <= maxColumn)
            {
                var targetTextBox = parentGrid.Children
                    .OfType<TextBox>()
                    .FirstOrDefault(textBox =>
                    {
                        int childRow = Grid.GetRow(textBox);
                        int childColumn = Grid.GetColumn(textBox);
                        int childRowSpan = Math.Max(1, Grid.GetRowSpan(textBox));
                        int childColumnSpan = Math.Max(1, Grid.GetColumnSpan(textBox));

                        return textBox != currentTextBox &&
                               nextRow >= childRow &&
                               nextRow < childRow + childRowSpan &&
                               nextColumn >= childColumn &&
                               nextColumn < childColumn + childColumnSpan;
                    });

                if (targetTextBox != null)
                {
                    return targetTextBox;
                }

                nextRow += rowStep;
                nextColumn += columnStep;
            }

            return null;
        }

        private bool ShouldMoveToAdjacentTableCell(TextBox currentTextBox, Key directionKey)
        {
            if (currentTextBox == null || Keyboard.Modifiers != ModifierKeys.None || currentTextBox.SelectionLength > 0)
            {
                return false;
            }

            switch (directionKey)
            {
                case Key.Left:
                    return currentTextBox.CaretIndex == 0;
                case Key.Right:
                    return currentTextBox.CaretIndex >= (currentTextBox.Text?.Length ?? 0);
                case Key.Up:
                    return GetCurrentTextBoxLineIndex(currentTextBox) <= 0;
                case Key.Down:
                    return GetCurrentTextBoxLineIndex(currentTextBox) >= GetLastTextBoxLineIndex(currentTextBox);
                default:
                    return false;
            }
        }

        private void MoveCaretToAdjacentTableCell(TextBox currentTextBox, TextBox targetTextBox, Key directionKey)
        {
            if (targetTextBox == null)
            {
                return;
            }

            targetTextBox.Focus();

            int targetCaretIndex;
            switch (directionKey)
            {
                case Key.Left:
                    targetCaretIndex = targetTextBox.Text?.Length ?? 0;
                    break;
                case Key.Right:
                    targetCaretIndex = 0;
                    break;
                default:
                    targetCaretIndex = Math.Min(currentTextBox?.CaretIndex ?? 0, targetTextBox.Text?.Length ?? 0);
                    break;
            }

            targetTextBox.CaretIndex = Math.Max(0, targetCaretIndex);
            targetTextBox.SelectionLength = 0;
        }

        private int GetCurrentTextBoxLineIndex(TextBox textBox)
        {
            if (textBox == null)
            {
                return 0;
            }

            return Math.Max(0, textBox.GetLineIndexFromCharacterIndex(textBox.CaretIndex));
        }

        private int GetLastTextBoxLineIndex(TextBox textBox)
        {
            if (textBox == null)
            {
                return 0;
            }

            return Math.Max(0, textBox.LineCount - 1);
        }

        private void AddTableRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

            InsertTableRow(context, context.Structure.HeaderRowCount + context.Structure.BodyRowCount + 1);
        }

        private void InsertTableRow(TableEditorContext context, int visualRow)
        {
            EnsureEditableTableStructure(context.Structure);

            if (visualRow <= context.Structure.HeaderRowCount)
            {
                ShiftCellsForInsertedRow(context.Structure.HeaderCells, visualRow);
                context.Structure.HeaderRowCount++;
            }
            else
            {
                int bodyRow = Math.Max(1, visualRow - context.Structure.HeaderRowCount);
                ShiftCellsForInsertedRow(context.Structure.BodyCells, bodyRow);
                context.Structure.BodyRowCount++;
            }

            RefreshTableEditor(context);
        }

        private void ShiftCellsForInsertedRow(List<TableCellDefinition> cells, int insertAtRow)
        {
            foreach (var cell in cells)
            {
                int cellEndRow = cell.Row + cell.RowSpan - 1;
                if (cell.Row >= insertAtRow)
                {
                    cell.Row++;
                }
                else if (cell.Row < insertAtRow && cellEndRow >= insertAtRow)
                {
                    cell.RowSpan++;
                }
            }
        }

        private string FormatTableHeaders(Table table)
        {
            var structure = ExtractTableStructure(table);
            return FormatTableHeaders(structure);
        }

        private string FormatTableHeaders(TableStructure structure)
        {
            return FormatTableCells(structure.HeaderCells);
        }

        private string FormatTableBodyRows(Table table)
        {
            var structure = ExtractTableStructure(table);
            return FormatTableBodyRows(structure);
        }

        private string FormatTableBodyRows(TableStructure structure)
        {
            return FormatTableCells(structure.BodyCells);
        }

        private string FormatTableCells(IEnumerable<TableCellDefinition> cells)
        {
            var orderedCells = cells
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToList();

            if (!orderedCells.Any())
            {
                return string.Empty;
            }

            var lines = orderedCells
                .GroupBy(cell => cell.Row)
                .OrderBy(group => group.Key)
                .Select(group => string.Join(" | ",
                    group.OrderBy(cell => cell.Column).Select(FormatTableCellToken)));

            return string.Join(Environment.NewLine, lines);
        }

        private async void SavePreviewTableButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

            if (await SaveTableAsync(context, false))
            {
                MessageBox.Show("Таблица сохранена.", "Таблица",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private async Task ExportNirPublishingTableByTemplate(TableEditorContext context)
        {
            try
            {
                NirPublishingImportTemplate? exportTemplate = ShowNirPublishingTemplateDialog(
                    "Экспорт из таблицы 3.1",
                    "Выберите шаблон экспорта для таблицы 3.1",
                    true);
                if (!exportTemplate.HasValue)
                {
                    return;
                }

                if (exportTemplate == NirPublishingImportTemplate.StudyPublishingPlan)
                {
                    await ExportToStudyPublishingPlan(context);
                    return;
                }

                await ExportToPublicationsListTemplate(context);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось экспортировать таблицу:{Environment.NewLine}{ex.Message}",
                    "Таблица",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async Task ExportToStudyPublishingPlan(TableEditorContext context)
        {
            var rows = BuildTable7ExportRows(context);
            if (!rows.Any())
            {
                MessageBox.Show(
                    "В таблице нет данных для экспорта в таблицу 7.",
                    "Таблица",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string fileName = $"publishPlan_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            string outputPath = Path.Combine(desktopPath, fileName);

            ExcelTableExchangeService.Export(outputPath, BuildStudyPublishingPlanExcelData(rows));

            MessageBox.Show(
                $"Файл сохранен:{Environment.NewLine}{outputPath}",
                "Таблица 7",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            await LogHistoryAsync(
                HistoryActionExport,
                "table",
                BuildTableHistoryLocation(context.Template, context.SubSection, context.Table),
                "Экспорт по шаблону \"План учебно-издательской деятельности\"");
        }

        private ExcelTableData BuildStudyPublishingPlanExcelData(IReadOnlyList<Table7ExportRow> rows)
        {
            var tableData = new ExcelTableData
            {
                ColumnCount = 11,
                HeaderRowCount = 2,
                BodyRowCount = rows?.Count ?? 0
            };

            double[] columnWidths =
            {
                6, 28, 18, 15, 10, 10, 9, 20, 24, 24, 14
            };
            for (int column = 1; column <= columnWidths.Length; column++)
            {
                tableData.ColumnWidths[column] = columnWidths[column - 1];
            }

            tableData.HeaderRowHeights[1] = 24;
            tableData.HeaderRowHeights[2] = 46;

            void AddHeaderCell(int row, int column, string text, int colSpan = 1, int rowSpan = 1, int? styleIndexOverride = null)
            {
                tableData.HeaderCells.Add(new ExcelTableCell
                {
                    Row = row,
                    Column = column,
                    Text = text ?? string.Empty,
                    ColSpan = colSpan,
                    RowSpan = rowSpan,
                    IsHeader = true,
                    StyleIndexOverride = styleIndexOverride
                });
            }

            void AddBodyCell(int row, int column, string text)
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    return;
                }

                tableData.BodyCells.Add(new ExcelTableCell
                {
                    Row = row,
                    Column = column,
                    Text = text ?? string.Empty,
                    IsHeader = false,
                    StyleIndexOverride = 0
                });
            }

            AddHeaderCell(1, 1, "7. План учебно-издательской деятельности", 11, 1, 2);
            AddHeaderCell(2, 1, "№", 1, 1, 1);
            AddHeaderCell(2, 2, "Наименование работ", 1, 1, 1);
            AddHeaderCell(2, 3, "Исполнители", 1, 1, 1);
            AddHeaderCell(2, 4, "Обоснование необходимости", 1, 1, 1);
            AddHeaderCell(2, 5, "Вид издания", 1, 1, 1);
            AddHeaderCell(2, 6, "Объем в уч.-изд. (листах)", 1, 1, 1);
            AddHeaderCell(2, 7, "Тираж", 1, 1, 1);
            AddHeaderCell(2, 8, "Наименование направления (специальности)", 1, 1, 1);
            AddHeaderCell(2, 9, "Дисциплина", 2, 1, 1);
            AddHeaderCell(2, 11, "Срок готовности", 1, 1, 1);

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                Table7ExportRow row = rows[rowIndex];
                int targetRow = rowIndex + 1;
                tableData.BodyRowHeights[targetRow] = 58;

                AddBodyCell(targetRow, 1, row.Number);
                AddBodyCell(targetRow, 2, row.WorkName);
                AddBodyCell(targetRow, 3, row.Performers);
                AddBodyCell(targetRow, 4, string.Empty);
                AddBodyCell(targetRow, 5, row.PublicationType);
                AddBodyCell(targetRow, 6, row.Volume);
                AddBodyCell(targetRow, 7, string.Empty);
                AddBodyCell(targetRow, 8, string.Empty);
                AddBodyCell(targetRow, 9, string.Empty);
                AddBodyCell(targetRow, 10, string.Empty);
                AddBodyCell(targetRow, 11, string.Empty);
            }

            return tableData;
        }

        private async Task ExportToPublicationsListTemplate(TableEditorContext context)
        {
            var rows = BuildNirPublicationExportRows(context);
            if (!rows.Any())
            {
                MessageBox.Show(
                    "В таблице нет данных для экспорта по шаблону \"Перечень публикаций сотрудников\".",
                    "Таблица",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string fileName = $"publicationsList_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            string outputPath = Path.Combine(desktopPath, fileName);

            ExcelTableExchangeService.Export(outputPath, BuildNirPublicationsTemplateExcelData(rows));

            MessageBox.Show(
                $"Файл сохранен:{Environment.NewLine}{outputPath}",
                "Перечень публикаций сотрудников",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            await LogHistoryAsync(
                HistoryActionExport,
                "table",
                BuildTableHistoryLocation(context.Template, context.SubSection, context.Table),
                "Экспорт по шаблону \"Перечень публикаций сотрудников\"");
        }

        private List<NirPublicationExportRow> BuildNirPublicationExportRows(TableEditorContext context)
        {
            EnsureEditableTableStructure(context.Structure);

            var result = new List<NirPublicationExportRow>();
            for (int bodyRow = 1; bodyRow <= Math.Max(0, context.Structure.BodyRowCount); bodyRow++)
            {
                string number = GetBodyCellText(context.Structure, bodyRow, 1);
                string share = GetBodyCellText(context.Structure, bodyRow, 2);
                string authors = GetBodyCellText(context.Structure, bodyRow, 3);
                string publicationName = GetBodyCellText(context.Structure, bodyRow, 4);
                string publicationType = GetBodyCellText(context.Structure, bodyRow, 5);
                string editionInfo = GetBodyCellText(context.Structure, bodyRow, 6);
                string publicationPlace = GetBodyCellText(context.Structure, bodyRow, 7);

                if (string.IsNullOrWhiteSpace(number) &&
                    string.IsNullOrWhiteSpace(share) &&
                    string.IsNullOrWhiteSpace(authors) &&
                    string.IsNullOrWhiteSpace(publicationName) &&
                    string.IsNullOrWhiteSpace(publicationType) &&
                    string.IsNullOrWhiteSpace(editionInfo) &&
                    string.IsNullOrWhiteSpace(publicationPlace))
                {
                    continue;
                }

                result.Add(new NirPublicationExportRow
                {
                    Number = number,
                    Share = share,
                    Authors = authors,
                    PublicationName = publicationName,
                    PublicationType = publicationType,
                    EditionInfo = editionInfo,
                    PublicationPlace = publicationPlace
                });
            }

            return result;
        }

        private ExcelTableData BuildNirPublicationsTemplateExcelData(IReadOnlyList<NirPublicationExportRow> rows)
        {
            var tableData = new ExcelTableData
            {
                ColumnCount = 17,
                HeaderRowCount = 5,
                BodyRowCount = rows?.Count ?? 0
            };

            double[] columnWidths =
            {
                18, 18, 18, 12, 18, 8, 12, 26, 44, 24, 42, 28, 10, 10, 10, 10, 10
            };
            for (int column = 1; column <= columnWidths.Length; column++)
            {
                tableData.ColumnWidths[column] = columnWidths[column - 1];
            }

            tableData.HeaderRowHeights[1] = 24;
            tableData.HeaderRowHeights[2] = 54;
            tableData.HeaderRowHeights[3] = 22;
            tableData.HeaderRowHeights[4] = 26;
            tableData.HeaderRowHeights[5] = 82;

            void AddHeaderCell(int row, int column, string text, int colSpan = 1, int rowSpan = 1, int? styleIndexOverride = null)
            {
                tableData.HeaderCells.Add(new ExcelTableCell
                {
                    Row = row,
                    Column = column,
                    Text = text ?? string.Empty,
                    ColSpan = colSpan,
                    RowSpan = rowSpan,
                    IsHeader = true,
                    StyleIndexOverride = styleIndexOverride
                });
            }

            void AddBodyCell(int row, int column, string text, int? styleIndexOverride = null)
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    return;
                }

                tableData.BodyCells.Add(new ExcelTableCell
                {
                    Row = row,
                    Column = column,
                    Text = text ?? string.Empty,
                    IsHeader = false,
                    StyleIndexOverride = styleIndexOverride
                });
            }

            AddHeaderCell(1, 1, "Таблица публикаций сотрудников СГУПС", 17, 1, 2);
            AddHeaderCell(
                2,
                1,
                "* ОБЯЗАТЕЛЬНОЕ ПОЛЕ 7! =1 - если соавторы СОТРУДНИКИ СГУПС с одной кафедры или соавторы не сотрудники СГУПС (доля не выделяется) = а<1 - если соавторы с разных кафедр, распределение должно быть согласованным, сумма долей по одной статье в отчетах разных кафедр должна быть равна 1",
                9,
                1,
                3);
            AddHeaderCell(3, 3, "Столбцы 6-12 копируются в таблицу 3.1 Отчета о НИР", 1, 1, 4);

            string[] numbering =
            {
                "1", "2", "3", "4", "5", "6", "7", "8", "9",
                "10", "11", "12", "13", "14", "15", "16", "17"
            };
            for (int column = 1; column <= numbering.Length; column++)
            {
                AddHeaderCell(4, column, numbering[column - 1], 1, 1, 1);
            }

            string[] headers =
            {
                "Факультет кафедры, заполняющей сведения",
                "факультет соавторов",
                "Наименование кафедры (заполняющей сведения)",
                "число кафедр",
                "Кафедра соавторов",
                "№",
                "доля авторов заполняющей кафедры* обязательное поле по согласованию с соавторами",
                "Фамилия И.О. авторов",
                "Наименование публикации",
                "Тип публикации",
                "Наименование издания, год, страницы публикации (без кавычек)",
                "Место издания (наименование организации, город)",
                "RSCI",
                "ВАК",
                "Scopus",
                "к1",
                "к2"
            };
            for (int column = 1; column <= headers.Length; column++)
            {
                AddHeaderCell(5, column, headers[column - 1], 1, 1, 1);
            }

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                NirPublicationExportRow row = rows[rowIndex];
                int targetRow = rowIndex + 1;
                tableData.BodyRowHeights[targetRow] = 60;

                AddBodyCell(targetRow, 1, string.Empty, 0);
                AddBodyCell(targetRow, 2, string.Empty, 0);
                AddBodyCell(targetRow, 3, string.Empty, 0);
                AddBodyCell(targetRow, 4, string.Empty, 0);
                AddBodyCell(targetRow, 5, string.Empty, 0);
                AddBodyCell(targetRow, 6, row.Number, 0);
                AddBodyCell(targetRow, 7, row.Share, 0);
                AddBodyCell(targetRow, 8, row.Authors, 0);
                AddBodyCell(targetRow, 9, row.PublicationName, 0);
                AddBodyCell(targetRow, 10, row.PublicationType, 0);
                AddBodyCell(targetRow, 11, row.EditionInfo, 0);
                AddBodyCell(targetRow, 12, row.PublicationPlace, 0);
                AddBodyCell(targetRow, 13, string.Empty, 0);
                AddBodyCell(targetRow, 14, string.Empty, 0);
                AddBodyCell(targetRow, 15, string.Empty, 0);
                AddBodyCell(targetRow, 16, string.Empty, 0);
                AddBodyCell(targetRow, 17, string.Empty, 0);
            }

            return tableData;
        }

        private List<Table7ExportRow> BuildTable7ExportRows(TableEditorContext context)
        {
            EnsureEditableTableStructure(context.Structure);

            var result = new List<Table7ExportRow>();
            for (int bodyRow = 1; bodyRow <= Math.Max(0, context.Structure.BodyRowCount); bodyRow++)
            {
                string number = GetBodyCellText(context.Structure, bodyRow, 1);
                string workName = GetBodyCellText(context.Structure, bodyRow, 4);
                string performers = GetBodyCellText(context.Structure, bodyRow, 3);
                string publicationType = GetBodyCellText(context.Structure, bodyRow, 5);
                string editionInfo = GetBodyCellText(context.Structure, bodyRow, 6);

                if (string.IsNullOrWhiteSpace(workName) &&
                    string.IsNullOrWhiteSpace(performers) &&
                    string.IsNullOrWhiteSpace(publicationType) &&
                    string.IsNullOrWhiteSpace(editionInfo))
                {
                    continue;
                }

                result.Add(new Table7ExportRow
                {
                    Number = number,
                    WorkName = workName,
                    Performers = performers,
                    PublicationType = publicationType,
                    Volume = CalculateTable7Volume(editionInfo)
                });
            }

            return result;
        }

        private string CalculateTable7Volume(string editionInfo)
        {
            string text = (editionInfo ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var match = Regex.Match(
                text,
                @"с\.\s*(\d+)\s*[-–—]\s*(\d+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!match.Success ||
                !int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int startPage) ||
                !int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int endPage) ||
                endPage < startPage)
            {
                return string.Empty;
            }

            decimal volume = Math.Round((endPage - startPage + 1) / 16m, 1, MidpointRounding.AwayFromZero);
            return volume.ToString("0.0", CultureInfo.GetCultureInfo("ru-RU"));
        }

        private string BuildTable7SpreadsheetXml(IReadOnlyList<Table7ExportRow> rows)
        {
            string Escape(string value) => System.Security.SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;

            void AppendCell(System.Text.StringBuilder xmlBuilder, int index, string styleId, string value, bool mergeDiscipline = false)
            {
                xmlBuilder.Append("<Cell ss:Index=\"")
                    .Append(index)
                    .Append("\" ss:StyleID=\"")
                    .Append(styleId)
                    .Append("\"");

                if (mergeDiscipline)
                {
                    xmlBuilder.Append(" ss:MergeAcross=\"1\"");
                }

                xmlBuilder.Append("><Data ss:Type=\"String\">")
                    .Append(Escape(value))
                    .Append("</Data></Cell>");
            }

            var builder = new System.Text.StringBuilder();
            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>")
                .Append("<?mso-application progid=\"Excel.Sheet\"?>")
                .Append("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" ")
                .Append("xmlns:o=\"urn:schemas-microsoft-com:office:office\" ")
                .Append("xmlns:x=\"urn:schemas-microsoft-com:office:excel\" ")
                .Append("xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" ")
                .Append("xmlns:html=\"http://www.w3.org/TR/REC-html40\">")
                .Append("<Styles>")
                .Append("<Style ss:ID=\"20\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\" ss:WrapText=\"1\"/><Borders><Border ss:Position=\"Bottom\"/><Border ss:Position=\"Top\"/><Border ss:Position=\"Left\"/><Border ss:Position=\"Right\"/><Border ss:Position=\"DiagonalLeft\"/></Borders><Font ss:FontName=\"Times New Roman\" ss:Size=\"12.0\" ss:Bold=\"1\" ss:Color=\"#000000\"/></Style>")
                .Append("<Style ss:ID=\"22\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\" ss:WrapText=\"1\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#000000\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#000000\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#000000\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#000000\"/><Border ss:Position=\"DiagonalLeft\"/></Borders><Font ss:FontName=\"Times New Roman\" ss:Size=\"12.0\" ss:Color=\"#000000\"/></Style>")
                .Append("<Style ss:ID=\"31\"><Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Top\" ss:WrapText=\"1\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#000000\"/><Border ss:Position=\"Top\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#000000\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#000000\"/><Border ss:Position=\"DiagonalLeft\"/></Borders><Font ss:FontName=\"sans-serif\" ss:Size=\"12.0\" ss:Color=\"#000000\"/></Style>")
                .Append("<Style ss:ID=\"32\"><Alignment ss:Horizontal=\"Left\" ss:Vertical=\"Top\" ss:WrapText=\"1\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#000000\"/><Border ss:Position=\"Top\"/><Border ss:Position=\"Left\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\" ss:Color=\"#000000\"/><Border ss:Position=\"DiagonalLeft\"/></Borders><Font ss:FontName=\"sans-serif\" ss:Size=\"12.0\" ss:Color=\"#000000\"/></Style>")
                .Append("</Styles>")
                .Append("<Worksheet ss:Name=\"Report\"><Table>")
                .Append("<Column ss:Width=\"21.632\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Column ss:Width=\"172.544\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Column ss:Width=\"105.728\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Column ss:Width=\"89.216\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Column ss:Width=\"54.784\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Column ss:Width=\"58.496\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Column ss:Width=\"51.84\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Column ss:Width=\"116.224\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Column ss:Width=\"135.424\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Column ss:Width=\"22.784\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Column ss:Width=\"73.472\" ss:AutoFitWidth=\"0\"/>")
                .Append("<Row ss:AutoFitHeight=\"1\"><Cell ss:Index=\"1\" ss:StyleID=\"20\" ss:MergeAcross=\"8\"><Data ss:Type=\"String\">7. План учебно-издательской деятельности</Data></Cell></Row>")
                .Append("<Row ss:AutoFitHeight=\"1\">");

            AppendCell(builder, 1, "22", "№");
            AppendCell(builder, 2, "22", "Наименование работ");
            AppendCell(builder, 3, "22", "Исполнители");
            AppendCell(builder, 4, "22", "Обоснование необходимости");
            AppendCell(builder, 5, "22", "Вид издания");
            AppendCell(builder, 6, "22", "Объем в уч.-изд. (листах)");
            AppendCell(builder, 7, "22", "Тираж");
            AppendCell(builder, 8, "22", "Наименование направления (специальности)");
            AppendCell(builder, 9, "22", "Дисциплина", true);
            AppendCell(builder, 11, "22", "Срок готовности");
            builder.Append("</Row>");

            foreach (var row in rows)
            {
                builder.Append("<Row ss:AutoFitHeight=\"1\">");
                AppendCell(builder, 1, "31", row.Number);
                AppendCell(builder, 2, "32", row.WorkName);
                AppendCell(builder, 3, "32", row.Performers);
                AppendCell(builder, 4, "32", string.Empty);
                AppendCell(builder, 5, "32", row.PublicationType);
                AppendCell(builder, 6, "32", row.Volume);
                AppendCell(builder, 7, "32", string.Empty);
                AppendCell(builder, 8, "32", string.Empty);
                AppendCell(builder, 9, "32", string.Empty, true);
                AppendCell(builder, 11, "32", string.Empty);
                builder.Append("</Row>");
            }

            builder.Append("</Table></Worksheet></Workbook>");
            return builder.ToString();
        }

        private async void ClearTableContentButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

            EnsureEditableTableStructure(context.Structure);

            if (context.Structure.BodyRowCount <= 0 && !context.Structure.BodyCells.Any())
            {
                return;
            }

            if (MessageBox.Show(
                "Очистить таблицу? Будут удалены все строки данных. Шапка и столбцы останутся без изменений.",
                "Таблица",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }

            context.Structure.BodyCells.Clear();
            context.Structure.BodyRowCount = 0;

            RefreshTableEditor(context);
            await SaveTableAsync(context, false);
        }

        private static bool CanImportNirPublishingTable(TableEditorContext context)
        {
            return string.Equals(context?.Table?.PatternName, NirPublishingTablePatternName, StringComparison.Ordinal);
        }

        private void ApplyImportedRowsToTable(TableEditorContext context, IReadOnlyList<string[]> importedRows)
        {
            if (context.Structure == null)
            {
                context.Structure = new TableStructure();
            }

            EnsureEditableTableStructure(context.Structure);
            context.Structure.ColumnCount = Math.Max(context.Structure.ColumnCount, 7);
            context.Structure.BodyCells.Clear();
            context.Structure.BodyRowCount = importedRows?.Count ?? 0;

            if (importedRows == null)
            {
                return;
            }

            for (int rowIndex = 0; rowIndex < importedRows.Count; rowIndex++)
            {
                string[] rowValues = importedRows[rowIndex] ?? Array.Empty<string>();
                for (int columnIndex = 0; columnIndex < 7; columnIndex++)
                {
                    context.Structure.BodyCells.Add(new TableCellDefinition
                    {
                        Row = rowIndex + 1,
                        Column = columnIndex + 1,
                        Text = columnIndex < rowValues.Length ? rowValues[columnIndex] ?? string.Empty : string.Empty,
                        IsHeader = false
                    });
                }
            }
        }

        private async Task ImportNirPublishingTableAsync(
            TableEditorContext context,
            string fileName,
            NirPublishingImportTemplate importTemplate)
        {
            var importedRows = importTemplate == NirPublishingImportTemplate.StudyPublishingPlan
                ? ExcelTableExchangeService.ImportNirPublishingPlanRows(fileName)
                : ExcelTableExchangeService.ImportNirPublishingRows(fileName);
            if (importedRows == null || importedRows.Count == 0)
            {
                MessageBox.Show(
                    importTemplate == NirPublishingImportTemplate.StudyPublishingPlan
                        ? "В выбранном файле не найдено строк для импорта из таблицы 7."
                        : "В выбранном Excel-файле не найдено строк для импорта в столбцах 6-12.",
                    "Таблица",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            ApplyImportedRowsToTable(context, importedRows);
            RefreshTableEditor(context);

            if (await SaveTableAsync(context, false))
            {
                MessageBox.Show(
                    "Таблица импортирована и сохранена.",
                    "Таблица",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);

                string importDetails = importTemplate == NirPublishingImportTemplate.StudyPublishingPlan
                    ? "Импорт по шаблону \"План учебно-издательской деятельности\""
                    : "Импорт по шаблону \"Перечень публикаций сотрудников\"";
                await LogHistoryAsync(
                    HistoryActionImport,
                    "table",
                    BuildTableHistoryLocation(context.Template, context.SubSection, context.Table),
                    importDetails);
            }
        }

        private NirPublishingImportTemplate? ShowNirPublishingTemplateDialog(
            string dialogTitle,
            string promptText,
            bool isExportAction)
        {
            NirPublishingImportTemplate? selectedTemplate = null;
            var dialog = new Window
            {
                Owner = this,
                Title = dialogTitle,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                SizeToContent = SizeToContent.WidthAndHeight,
                ShowInTaskbar = false,
                Background = Brushes.White,
                WindowStyle = WindowStyle.ToolWindow
            };

            var panel = new StackPanel
            {
                Margin = new Thickness(18),
                Width = 420
            };

            panel.Children.Add(new TextBlock
            {
                Text = promptText,
                FontWeight = FontWeights.SemiBold,
                FontSize = 16,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4")),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            panel.Children.Add(new TextBlock
            {
                Text = isExportAction
                    ? "1. Экспортировать по шаблону \"Перечень публикаций сотрудников\"\n2. Экспортировать по шаблону \"План учебно-издательской деятельности\""
                    : "1. Импортировать по шаблону \"Перечень публикаций сотрудников\"\n2. Импортировать по шаблону \"План учебно-издательской деятельности\"",
                Foreground = Brushes.DimGray,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 14)
            });

            var publicationsButton = CreateActionButton(
                isExportAction
                    ? "1. Экспортировать: Перечень публикаций сотрудников"
                    : "1. Импортировать: Перечень публикаций сотрудников");
            publicationsButton.Margin = new Thickness(0, 0, 0, 10);
            publicationsButton.Click += (sender, args) =>
            {
                selectedTemplate = NirPublishingImportTemplate.PublicationsList;
                dialog.DialogResult = true;
            };
            panel.Children.Add(publicationsButton);

            var planButton = CreateSecondaryButton(
                isExportAction
                    ? "2. Экспортировать: План учебно-издательской деятельности"
                    : "2. Импортировать: План учебно-издательской деятельности");
            planButton.Margin = new Thickness(0, 0, 0, 10);
            planButton.Click += (sender, args) =>
            {
                selectedTemplate = NirPublishingImportTemplate.StudyPublishingPlan;
                dialog.DialogResult = true;
            };
            panel.Children.Add(planButton);

            var cancelButton = CreateSecondaryButton("Отмена");
            cancelButton.Margin = new Thickness(0);
            cancelButton.Click += (sender, args) =>
            {
                dialog.DialogResult = false;
            };
            panel.Children.Add(cancelButton);

            dialog.Content = panel;
            return dialog.ShowDialog() == true ? selectedTemplate : null;
        }

        private async void ImportTableFromExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

            try
            {
                if (CanImportNirPublishingTable(context))
                {
                    NirPublishingImportTemplate? importTemplate = ShowNirPublishingTemplateDialog(
                        "Импорт в таблицу 3.1",
                        "Выберите шаблон импорта для таблицы 3.1",
                        false);
                    if (!importTemplate.HasValue)
                    {
                        return;
                    }

                    var nirOpenDialog = new OpenFileDialog
                    {
                        Title = importTemplate == NirPublishingImportTemplate.StudyPublishingPlan
                            ? "Импортировать таблицу 7"
                            : "Импортировать перечень публикаций сотрудников",
                        Filter = importTemplate == NirPublishingImportTemplate.StudyPublishingPlan
                            ? "Файлы Excel (*.xls;*.xlsx;*.xml)|*.xls;*.xlsx;*.xml|Excel 2003 XML (*.xls)|*.xls|Excel Workbook (*.xlsx)|*.xlsx|XML files (*.xml)|*.xml"
                            : "Excel Workbook (*.xlsx)|*.xlsx",
                        CheckFileExists = true,
                        Multiselect = false
                    };

                    if (nirOpenDialog.ShowDialog() != true)
                    {
                        return;
                    }

                    await ImportNirPublishingTableAsync(context, nirOpenDialog.FileName, importTemplate.Value);
                    return;
                }

                var openDialog = new OpenFileDialog
                {
                    Title = "Импортировать таблицу",
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    CheckFileExists = true,
                    Multiselect = false
                };

                if (openDialog.ShowDialog() != true)
                {
                    return;
                }

                var imported = ExcelTableExchangeService.Import(openDialog.FileName);
                context.Structure = ConvertFromExcelTableData(imported);

                if (context.TitleTextBox != null && string.IsNullOrWhiteSpace(context.TitleTextBox.Text))
                {
                    context.TitleTextBox.Text = Path.GetFileNameWithoutExtension(openDialog.FileName);
                }

                RefreshTableEditor(context);
                await LogHistoryAsync(
                    HistoryActionImport,
                    "table",
                    BuildTableHistoryLocation(context.Template, context.SubSection, context.Table),
                    "Импорт таблицы из Excel в редактор");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось импортировать таблицу:{Environment.NewLine}{ex.Message}",
                    "Таблица",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async Task<bool> SaveTableAsync(TableEditorContext context, bool refreshAfterSave)
        {
            if (context == null)
            {
                return false;
            }

            string title = (context.TitleTextBox?.Text ?? context.Table?.Title ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(title))
            {
                MessageBox.Show("Название таблицы не может быть пустым.", "Таблица",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            string previousTitle = context.Table.Title?.Trim() ?? string.Empty;
            string previousHeaders = FormatTableHeaders(context.Table);
            string previousBody = FormatTableBodyRows(context.Table);

            EnsureEditableTableStructure(context.Structure);
            if (!HasTableContent(context.Structure))
            {
                MessageBox.Show("Таблица не может быть пустой. Заполните хотя бы одну ячейку.", "Таблица",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            var database = new DataBase(context.Template.DatabasePath);
            context.Table.Title = title;
            await database.UpdateTable(context.Table);
            await database.DeleteAllTableItems(context.Table.Id);

            var structure = context.Structure;

            foreach (var headerCell in structure.HeaderCells)
            {
                await database.AddTableItem(new TableItem
                {
                    TableId = context.Table.Id,
                    Row = headerCell.Row,
                    Column = headerCell.Column,
                    Header = headerCell.Text,
                    ColSpan = headerCell.ColSpan,
                    RowSpan = headerCell.RowSpan,
                    IsHeader = headerCell.IsHeader
                });
            }

            foreach (var bodyCell in structure.BodyCells)
            {
                await database.AddTableItem(new TableItem
                {
                    TableId = context.Table.Id,
                    Row = structure.HeaderRowCount + bodyCell.Row,
                    Column = bodyCell.Column,
                    Header = bodyCell.Text,
                    ColSpan = bodyCell.ColSpan,
                    RowSpan = bodyCell.RowSpan,
                    IsHeader = bodyCell.IsHeader
                });
            }

            context.Table.TableItems = BuildTableItemsFromStructure(context.Table.Id, structure);
            string newHeaders = FormatTableHeaders(structure);
            string newBody = FormatTableBodyRows(structure);

            bool tableChanged =
                !string.Equals(previousTitle, title, StringComparison.Ordinal) ||
                !string.Equals(previousHeaders, newHeaders, StringComparison.Ordinal) ||
                !string.Equals(previousBody, newBody, StringComparison.Ordinal);

            if (tableChanged)
            {
                await LogHistoryAsync(
                    HistoryActionEdit,
                    "table",
                    BuildTableHistoryLocation(context.Template, context.SubSection, context.Table),
                    BuildHistoryDetails(
                        !string.Equals(previousTitle, title, StringComparison.Ordinal) ? "обновлено название" : null,
                        !string.Equals(previousHeaders, newHeaders, StringComparison.Ordinal) ? "обновлена шапка" : null,
                        !string.Equals(previousBody, newBody, StringComparison.Ordinal) ? "обновлены данные" : null));
            }

            if (refreshAfterSave)
            {
                await RefreshTemplateEntryAsync(context.Template, subsectionId: context.SubSection.Id);
            }

            return true;
        }

        private TableStructure ConvertFromExcelTableData(ExcelTableData data)
        {
            var structure = new TableStructure
            {
                ColumnCount = Math.Max(1, data?.ColumnCount ?? 1),
                HeaderRowCount = Math.Max(1, data?.HeaderRowCount ?? 1),
                BodyRowCount = Math.Max(0, data?.BodyRowCount ?? 0)
            };

            if (data != null)
            {
                foreach (var cell in data.HeaderCells)
                {
                    structure.HeaderCells.Add(new TableCellDefinition
                    {
                        Text = cell.Text ?? string.Empty,
                        Column = cell.Column,
                        Row = cell.Row,
                        ColSpan = Math.Max(1, cell.ColSpan),
                        RowSpan = Math.Max(1, cell.RowSpan),
                        IsHeader = cell.IsHeader
                    });
                }

                foreach (var cell in data.BodyCells)
                {
                    structure.BodyCells.Add(new TableCellDefinition
                    {
                        Text = cell.Text ?? string.Empty,
                        Column = cell.Column,
                        Row = cell.Row,
                        ColSpan = Math.Max(1, cell.ColSpan),
                        RowSpan = Math.Max(1, cell.RowSpan),
                        IsHeader = cell.IsHeader
                    });
                }
            }

            EnsureEditableTableStructure(structure);
            return structure;
        }

        private bool HasTableContent(TableStructure structure)
        {
            if (structure == null)
            {
                return false;
            }

            return structure.HeaderCells.Any(cell => !string.IsNullOrWhiteSpace(cell.Text))
                || structure.BodyCells.Any(cell => !string.IsNullOrWhiteSpace(cell.Text));
        }

        private List<TableItem> BuildTableItemsFromStructure(int tableId, TableStructure structure)
        {
            var items = new List<TableItem>();

            foreach (var headerCell in structure.HeaderCells)
            {
                items.Add(new TableItem
                {
                    TableId = tableId,
                    Row = headerCell.Row,
                    Column = headerCell.Column,
                    Header = headerCell.Text,
                    ColSpan = headerCell.ColSpan,
                    RowSpan = headerCell.RowSpan,
                    IsHeader = headerCell.IsHeader
                });
            }

            foreach (var bodyCell in structure.BodyCells)
            {
                items.Add(new TableItem
                {
                    TableId = tableId,
                    Row = structure.HeaderRowCount + bodyCell.Row,
                    Column = bodyCell.Column,
                    Header = bodyCell.Text,
                    ColSpan = bodyCell.ColSpan,
                    RowSpan = bodyCell.RowSpan,
                    IsHeader = bodyCell.IsHeader
                });
            }

            return items;
        }

        private Border CreateTextCard(string title, string text)
        {
            return CreateContentCard(title, new TextBlock
            {
                Text = text,
                TextWrapping = TextWrapping.Wrap,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#343a40"))
            });
        }

        private Border CreateContentCard(string title, UIElement content)
        {
            var panel = new StackPanel();
            panel.Children.Add(new TextBlock
            {
                Text = title,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4"))
            });
            panel.Children.Add(content);

            return new Border
            {
                Background = Brushes.White,
                CornerRadius = new CornerRadius(14),
                Padding = new Thickness(16),
                Margin = new Thickness(0, 0, 0, 12),
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#d6e2ea")),
                BorderThickness = new Thickness(1),
                Child = panel
            };
        }

        private UIElement CreateTablePreview(Table table)
        {
            return CreateStyledTablePreview(ExtractTableStructure(table));
        }

        private sealed class TableCellDefinition
        {
            public string Text { get; set; }
            public int Column { get; set; }
            public int Row { get; set; }
            public int ColSpan { get; set; } = 1;
            public int RowSpan { get; set; } = 1;
            public bool IsHeader { get; set; }
        }

        private sealed class TableStructure
        {
            public List<TableCellDefinition> HeaderCells { get; } = new List<TableCellDefinition>();
            public List<TableCellDefinition> BodyCells { get; } = new List<TableCellDefinition>();
            public int ColumnCount { get; set; }
            public int HeaderRowCount { get; set; }
            public int BodyRowCount { get; set; }
        }

        private const double PreviewTableColumnWidth = 160;

        private TableStructure ExtractTableStructure(Table table)
        {
            var structure = new TableStructure();
            if (table?.TableItems == null || !table.TableItems.Any())
            {
                return structure;
            }

            var items = table.TableItems
                .OrderBy(item => item.Row)
                .ThenBy(item => item.Column)
                .ToList();

            bool hasExplicitHeaderItems = items.Any(item => item.IsHeader);
            if (hasExplicitHeaderItems)
            {
                bool isLegacyExplicitLayout = items.Any(item => !item.IsHeader && item.Row == 1);
                if (isLegacyExplicitLayout)
                {
                    foreach (var headerItem in items.Where(item => item.IsHeader))
                    {
                        structure.HeaderCells.Add(new TableCellDefinition
                        {
                            Text = headerItem.Header ?? string.Empty,
                            Column = headerItem.Column,
                            Row = headerItem.Row,
                            ColSpan = Math.Max(1, headerItem.ColSpan),
                            RowSpan = Math.Max(1, headerItem.RowSpan),
                            IsHeader = true
                        });
                    }

                    foreach (var bodyItem in items.Where(item => !item.IsHeader))
                    {
                        structure.BodyCells.Add(new TableCellDefinition
                        {
                            Text = bodyItem.Header ?? string.Empty,
                            Column = bodyItem.Column,
                            Row = bodyItem.Row,
                            ColSpan = Math.Max(1, bodyItem.ColSpan),
                            RowSpan = Math.Max(1, bodyItem.RowSpan),
                            IsHeader = false
                        });
                    }
                }
                else
                {
                    int headerRowCount = Math.Max(1, CountLeadingHeaderRows(items));

                    foreach (var item in items)
                    {
                        var cell = new TableCellDefinition
                        {
                            Text = item.Header ?? string.Empty,
                            Column = item.Column,
                            Row = item.Row,
                            ColSpan = Math.Max(1, item.ColSpan),
                            RowSpan = Math.Max(1, item.RowSpan),
                            IsHeader = item.IsHeader
                        };

                        if (item.Row <= headerRowCount)
                        {
                            structure.HeaderCells.Add(cell);
                        }
                        else
                        {
                            cell.Row = item.Row - headerRowCount;
                            structure.BodyCells.Add(cell);
                        }
                    }
                }
            }
            else
            {
                var groupedRows = items
                    .GroupBy(item => item.Row)
                    .OrderBy(group => group.Key)
                    .Select(group => group.OrderBy(item => item.Column).ToList())
                    .ToList();

                if (groupedRows.Any())
                {
                    int columnIndex = 1;
                    foreach (var item in groupedRows[0])
                    {
                        structure.HeaderCells.Add(new TableCellDefinition
                        {
                            Text = item.Header ?? string.Empty,
                            Column = columnIndex++,
                            Row = 1,
                            ColSpan = 1,
                            RowSpan = 1,
                            IsHeader = true
                        });
                    }
                }

                int bodyRowIndex = 1;
                foreach (var bodyRow in groupedRows.Skip(1))
                {
                    foreach (var bodyItem in bodyRow)
                    {
                        structure.BodyCells.Add(new TableCellDefinition
                        {
                            Text = bodyItem.Header ?? string.Empty,
                            Column = bodyItem.Column,
                            Row = bodyRowIndex,
                            ColSpan = Math.Max(1, bodyItem.ColSpan),
                            RowSpan = Math.Max(1, bodyItem.RowSpan),
                            IsHeader = false
                        });
                    }

                    bodyRowIndex++;
                }
            }

            return NormalizeTableStructure(structure);
        }

        private int CountLeadingHeaderRows(IReadOnlyCollection<TableItem> items)
        {
            if (items == null || items.Count == 0)
            {
                return 0;
            }

            int maxRow = items.Max(item => item.Row + Math.Max(1, item.RowSpan) - 1);
            int headerRows = 0;

            for (int row = 1; row <= maxRow; row++)
            {
                var rowItems = items
                    .Where(item => item.Row <= row && row < item.Row + Math.Max(1, item.RowSpan))
                    .ToList();

                if (!rowItems.Any() || rowItems.Any(item => !item.IsHeader))
                {
                    break;
                }

                headerRows = row;
            }

            return headerRows;
        }

        private TableStructure ParseTableInput(string headersText, string rowsText)
        {
            var structure = new TableStructure();
            var headerLines = (headersText ?? string.Empty)
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            var occupied = new HashSet<string>(StringComparer.Ordinal);
            for (int rowIndex = 0; rowIndex < headerLines.Length; rowIndex++)
            {
                int currentColumn = 1;
                foreach (var rawCell in SplitTableLine(headerLines[rowIndex]))
                {
                    while (occupied.Contains($"{rowIndex + 1}:{currentColumn}"))
                    {
                        currentColumn++;
                    }

                    var cell = ParseTableCell(rawCell, rowIndex + 1, currentColumn, true);
                    structure.HeaderCells.Add(cell);

                    for (int row = cell.Row; row < cell.Row + cell.RowSpan; row++)
                    {
                        for (int column = cell.Column; column < cell.Column + cell.ColSpan; column++)
                        {
                            occupied.Add($"{row}:{column}");
                        }
                    }

                    currentColumn += cell.ColSpan;
                }
            }

            var lines = (rowsText ?? string.Empty)
                .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                int rowIndex = structure.BodyRowCount + 1;
                int currentColumn = 1;
                foreach (var rawCell in SplitTableLine(line))
                {
                    while (occupied.Contains($"body:{rowIndex}:{currentColumn}"))
                    {
                        currentColumn++;
                    }

                    var cell = ParseTableCell(rawCell, rowIndex, currentColumn, false);
                    structure.BodyCells.Add(cell);

                    for (int row = cell.Row; row < cell.Row + cell.RowSpan; row++)
                    {
                        for (int column = cell.Column; column < cell.Column + cell.ColSpan; column++)
                        {
                            occupied.Add($"body:{row}:{column}");
                        }
                    }

                    currentColumn += cell.ColSpan;
                }

                structure.BodyRowCount = Math.Max(structure.BodyRowCount, rowIndex);
            }

            return NormalizeTableStructure(structure);
        }

        private List<string> SplitTableLine(string text)
        {
            return (text ?? string.Empty)
                .Split(new[] { '|' }, StringSplitOptions.None)
                .Select(cell => cell.Trim())
                .ToList();
        }

        private TableStructure NormalizeTableStructure(TableStructure structure)
        {
            int headerColumnCount = structure.HeaderCells.Any()
                ? structure.HeaderCells.Max(cell => cell.Column + cell.ColSpan - 1)
                : 0;
            int bodyColumnCount = structure.BodyCells.Any()
                ? structure.BodyCells.Max(cell => cell.Column + cell.ColSpan - 1)
                : 0;

            int headerRowCount = structure.HeaderCells.Any()
                ? structure.HeaderCells.Max(cell => cell.Row + cell.RowSpan - 1)
                : 0;
            int bodyRowCount = structure.BodyCells.Any()
                ? structure.BodyCells.Max(cell => cell.Row + cell.RowSpan - 1)
                : 0;

            structure.ColumnCount = Math.Max(structure.ColumnCount, Math.Max(headerColumnCount, bodyColumnCount));
            structure.HeaderRowCount = Math.Max(structure.HeaderRowCount, headerRowCount);
            structure.BodyRowCount = Math.Max(structure.BodyRowCount, bodyRowCount);

            return structure;
        }

        private UIElement CreateStyledTablePreview(TableStructure structure)
        {
            int columnCount = structure.ColumnCount;

            if (columnCount == 0)
            {
                return new Border
                {
                    Background = Brushes.White,
                    BorderBrush = Brushes.Black,
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(12),
                    Child = new TextBlock
                    {
                        Text = "Таблица пока пустая.",
                        Foreground = Brushes.Gray
                    }
                };
            }

            var panel = new StackPanel
            {
                Background = Brushes.White
            };

            if (structure.HeaderRowCount > 0)
            {
                var headerGrid = new Grid();
                ConfigurePreviewTableColumns(headerGrid, columnCount);

                for (int rowIndex = 0; rowIndex < structure.HeaderRowCount; rowIndex++)
                {
                    headerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                }

                AddHeaderCells(headerGrid, structure, columnCount);
                panel.Children.Add(headerGrid);
            }

            if (structure.BodyRowCount > 0)
            {
                var bodyGrid = new Grid();
                ConfigurePreviewTableColumns(bodyGrid, columnCount);

                for (int rowIndex = 0; rowIndex < structure.BodyRowCount; rowIndex++)
                {
                    bodyGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                }

                AddTableCells(bodyGrid, structure.BodyCells, structure.BodyRowCount, columnCount, false);
                panel.Children.Add(bodyGrid);
            }

            return new Border
            {
                Background = Brushes.White,
                BorderBrush = Brushes.Black,
                BorderThickness = new Thickness(1),
                ClipToBounds = true,
                Child = new ScrollViewer
                {
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    Content = panel
                }
            };
        }

        private void ConfigurePreviewTableColumns(Grid grid, int columnCount)
        {
            grid.ClipToBounds = true;

            for (int index = 0; index < columnCount; index++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition
                {
                    Width = new GridLength(PreviewTableColumnWidth)
                });
            }
        }

        private void AddHeaderCells(Grid root, TableStructure structure, int columnCount)
        {
            AddTableCells(root, structure.HeaderCells, structure.HeaderRowCount, columnCount, true);
        }

        private void AddTableCells(Grid root, IReadOnlyCollection<TableCellDefinition> cells, int rowCount, int columnCount, bool isHeader)
        {
            var occupied = new bool[rowCount + 1, columnCount + 1];

            foreach (var cell in cells)
            {
                var border = CreateTableCellBorder(cell.Text, cell.IsHeader);
                Grid.SetRow(border, cell.Row - 1);
                Grid.SetColumn(border, cell.Column - 1);
                Grid.SetColumnSpan(border, cell.ColSpan);
                Grid.SetRowSpan(border, cell.RowSpan);
                root.Children.Add(border);

                for (int row = cell.Row; row < cell.Row + cell.RowSpan; row++)
                {
                    for (int column = cell.Column; column < cell.Column + cell.ColSpan; column++)
                    {
                        occupied[row, column] = true;
                    }
                }
            }

            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    if (occupied[row, column])
                    {
                        continue;
                    }

                    var emptyBorder = CreateTableCellBorder(string.Empty, isHeader);
                    Grid.SetRow(emptyBorder, row - 1);
                    Grid.SetColumn(emptyBorder, column - 1);
                    root.Children.Add(emptyBorder);
                }
            }
        }

        private Grid CreateStyledTableRow(IReadOnlyList<string> cells, int columnCount, bool isHeader)
        {
            var grid = new Grid
            {
                Background = Brushes.White,
                ClipToBounds = true
            };

            ConfigurePreviewTableColumns(grid, columnCount);

            for (int index = 0; index < columnCount; index++)
            {
                var border = CreateTableCellBorder(
                    index < cells.Count ? cells[index] : string.Empty,
                    isHeader);

                Grid.SetColumn(border, index);
                grid.Children.Add(border);
            }

            return grid;
        }

        private Border CreateTableCellBorder(string text, bool isHeader)
        {
            return new Border
            {
                BorderBrush = Brushes.Black,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Padding = new Thickness(8, 10, 8, 10),
                Background = Brushes.White,
                ClipToBounds = true,
                Child = new TextBlock
                {
                    Text = text,
                    FontSize = isHeader ? 10 : 12,
                    FontWeight = isHeader ? FontWeights.Bold : FontWeights.Normal,
                    Foreground = Brushes.Black,
                    TextAlignment = TextAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextWrapping = TextWrapping.Wrap,
                    TextTrimming = TextTrimming.None
                }
            };
        }

        private TableCellDefinition ParseTableCell(string rawValue, int row, int column, bool isHeader)
        {
            string value = rawValue?.Trim() ?? string.Empty;
            int colSpan = 1;
            int rowSpan = 1;
            string text = value;

            var match = Regex.Match(value, @"^(?<text>.*?)(?:\[(?<spec>[^\]]+)\])?$");
            if (match.Success)
            {
                text = match.Groups["text"].Value.Trim();
                string spec = match.Groups["spec"].Value;
                if (!string.IsNullOrWhiteSpace(spec))
                {
                    foreach (var part in spec.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        string trimmed = part.Trim().ToLowerInvariant();
                        if (trimmed.StartsWith("c") && int.TryParse(trimmed.Substring(1), out int parsedColSpan))
                        {
                            colSpan = Math.Max(1, parsedColSpan);
                        }
                        else if (trimmed.StartsWith("r") && int.TryParse(trimmed.Substring(1), out int parsedRowSpan))
                        {
                            rowSpan = Math.Max(1, parsedRowSpan);
                        }
                    }
                }
            }

            return new TableCellDefinition
            {
                Text = text,
                Row = row,
                Column = column,
                ColSpan = colSpan,
                RowSpan = rowSpan,
                IsHeader = isHeader
            };
        }

        private string FormatTableCellToken(TableCellDefinition cell)
        {
            var spans = new List<string>();
            if (cell.ColSpan > 1)
            {
                spans.Add($"c{cell.ColSpan}");
            }

            if (cell.RowSpan > 1)
            {
                spans.Add($"r{cell.RowSpan}");
            }

            return spans.Count == 0
                ? cell.Text
                : $"{cell.Text}[{string.Join(",", spans)}]";
        }

        public class CalendarDay
        {
            public string Day { get; set; }
            public string TextColor { get; set; }
            public string FontWeight { get; set; }
            public string BackgroundColor { get; set; }
            public DateTime? Date { get; set; }
        }

        private void InitializeDateRange()
        {
            var today = DateTime.Today;
            var startDate = new DateTime(today.Year, 1, 1);
            var endDate = new DateTime(today.Year, 12, 31);

            UpdateDateDisplay(startDate, endDate);
            InitializeDateComboBoxes();
            SetDateInComboBoxes(startDate, endDate);
            UpdateCalendarDisplay(today);
        }

        private void UpdateHistoryCalendarVisibility(bool isVisible)
        {
            if (HistoryPeriodPanel != null)
            {
                HistoryPeriodPanel.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
            }

            if (!isVisible && CalendarPopup != null)
            {
                CalendarPopup.IsOpen = false;
            }
        }

        private void InitializeDateComboBoxes()
        {
            for (int i = 1; i <= 31; i++)
            {
                StartDayComboBox.Items.Add(i);
                EndDayComboBox.Items.Add(i);
            }

            var months = new[]
            {
                "Янв", "Фев", "Мар", "Апр", "Май", "Июн",
                "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"
            };

            for (int i = 0; i < months.Length; i++)
            {
                StartMonthComboBox.Items.Add(new { Text = months[i], Value = i + 1 });
                EndMonthComboBox.Items.Add(new { Text = months[i], Value = i + 1 });
            }

            int currentYear = DateTime.Today.Year;
            for (int i = currentYear - 5; i <= currentYear + 5; i++)
            {
                StartYearComboBox.Items.Add(i);
                EndYearComboBox.Items.Add(i);
            }

            StartMonthComboBox.DisplayMemberPath = "Text";
            EndMonthComboBox.DisplayMemberPath = "Text";
        }

        private void StartDate_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var startDate = GetStartDateFromComboBoxes();
            if (startDate.HasValue)
            {
                UpdateCalendarDisplay(_currentCalendarDate);
                var endDate = GetEndDateFromComboBoxes();
                if (!endDate.HasValue)
                {
                    SetDateInComboBoxes(startDate.Value, startDate.Value);
                }
            }
        }

        private void EndDate_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var endDate = GetEndDateFromComboBoxes();
            if (endDate.HasValue)
            {
                UpdateCalendarDisplay(_currentCalendarDate);
                var startDate = GetStartDateFromComboBoxes();
                if (!startDate.HasValue)
                {
                    SetDateInComboBoxes(endDate.Value, endDate.Value);
                }
            }
        }

        private DateTime? GetDateFromComboBoxes(ComboBox dayCombo, ComboBox monthCombo, ComboBox yearCombo)
        {
            if (dayCombo.SelectedItem != null && monthCombo.SelectedItem != null && yearCombo.SelectedItem != null)
            {
                try
                {
                    int day = (int)dayCombo.SelectedItem;
                    int month = ((dynamic)monthCombo.SelectedItem).Value;
                    int year = (int)yearCombo.SelectedItem;

                    if (DateTime.DaysInMonth(year, month) >= day)
                    {
                        return new DateTime(year, month, day);
                    }
                }
                catch
                {
                    return null;
                }
            }
            return null;
        }

        private DateTime? GetStartDateFromComboBoxes()
        {
            return GetDateFromComboBoxes(StartDayComboBox, StartMonthComboBox, StartYearComboBox);
        }

        private DateTime? GetEndDateFromComboBoxes()
        {
            return GetDateFromComboBoxes(EndDayComboBox, EndMonthComboBox, EndYearComboBox);
        }

        private void SetDateInComboBoxes(DateTime startDate, DateTime endDate)
        {
            StartDayComboBox.SelectedItem = startDate.Day;
            StartMonthComboBox.SelectedIndex = startDate.Month - 1;
            StartYearComboBox.SelectedItem = startDate.Year;

            EndDayComboBox.SelectedItem = endDate.Day;
            EndMonthComboBox.SelectedIndex = endDate.Month - 1;
            EndYearComboBox.SelectedItem = endDate.Year;
        }

        private void UpdateCalendarDisplay(DateTime date)
        {
            _currentCalendarDate = date;
            CurrentMonthYear.Text = date.ToString("MMMM yyyy", new CultureInfo("ru-RU"));
            GenerateCalendarDays(date);
        }

        private void GenerateCalendarDays(DateTime date)
        {
            var calendarDays = new List<CalendarDay>();
            var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
            int firstDayOfWeek = ((int)firstDayOfMonth.DayOfWeek == 0) ? 7 : (int)firstDayOfMonth.DayOfWeek;

            var previousMonth = firstDayOfMonth.AddMonths(-1);
            var daysInPreviousMonth = DateTime.DaysInMonth(previousMonth.Year, previousMonth.Month);

            for (int i = daysInPreviousMonth - firstDayOfWeek + 2; i <= daysInPreviousMonth; i++)
            {
                calendarDays.Add(new CalendarDay
                {
                    Day = i.ToString(),
                    TextColor = "#adb5bd",
                    FontWeight = "Normal",
                    BackgroundColor = "Transparent"
                });
            }

            for (int i = 1; i <= lastDayOfMonth.Day; i++)
            {
                var currentDay = new DateTime(date.Year, date.Month, i);
                bool isSelected = IsDateInSelectedRange(currentDay);
                bool isToday = currentDay.Date == DateTime.Today;

                calendarDays.Add(new CalendarDay
                {
                    Day = i.ToString(),
                    TextColor = isSelected ? "White" : (isToday ? "#0167a4" : "#495057"),
                    FontWeight = isToday ? "Bold" : "Normal",
                    BackgroundColor = isSelected ? "#0167a4" : "Transparent",
                    Date = currentDay
                });
            }

            int totalCells = 42;
            int nextMonthDay = 1;
            while (calendarDays.Count < totalCells)
            {
                calendarDays.Add(new CalendarDay
                {
                    Day = nextMonthDay.ToString(),
                    TextColor = "#adb5bd",
                    FontWeight = "Normal",
                    BackgroundColor = "Transparent"
                });
                nextMonthDay++;
            }

            CalendarDays.ItemsSource = calendarDays;
        }

        private bool IsDateInSelectedRange(DateTime date)
        {
            var startDate = GetStartDateFromComboBoxes();
            var endDate = GetEndDateFromComboBoxes();

            if (startDate.HasValue && endDate.HasValue)
            {
                return date.Date >= startDate.Value.Date && date.Date <= endDate.Value.Date;
            }
            else if (startDate.HasValue && !endDate.HasValue)
            {
                return date.Date == startDate.Value.Date;
            }

            return false;
        }

        private void CalendarDay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.DataContext is CalendarDay day && day.Date.HasValue)
            {
                var clickedDate = day.Date.Value;

                if (clickedDate.Month != _currentCalendarDate.Month || clickedDate.Year != _currentCalendarDate.Year)
                {
                    _currentCalendarDate = new DateTime(clickedDate.Year, clickedDate.Month, 1);
                }

                var currentStartDate = GetStartDateFromComboBoxes();
                var currentEndDate = GetEndDateFromComboBoxes();

                if (!currentStartDate.HasValue || !currentEndDate.HasValue ||
                    (currentStartDate.HasValue && currentEndDate.HasValue))
                {
                    SetDateInComboBoxes(clickedDate, clickedDate);
                }
                else if (currentStartDate.HasValue && !currentEndDate.HasValue)
                {
                    if (clickedDate < currentStartDate.Value)
                    {
                        SetDateInComboBoxes(clickedDate, clickedDate);
                    }
                    else
                    {
                        SetDateInComboBoxes(currentStartDate.Value, clickedDate);
                    }
                }

                UpdateCalendarDisplay(_currentCalendarDate);
            }
        }

        private void PreviousMonth_Click(object sender, RoutedEventArgs e)
        {
            _currentCalendarDate = _currentCalendarDate.AddMonths(-1);
            UpdateCalendarDisplay(_currentCalendarDate);
        }

        private void NextMonth_Click(object sender, RoutedEventArgs e)
        {
            _currentCalendarDate = _currentCalendarDate.AddMonths(1);
            UpdateCalendarDisplay(_currentCalendarDate);
        }

        private void PeriodSelector_Click(object sender, RoutedEventArgs e)
        {
            if (!(MainContentControl.Content is HistoryChangesView))
            {
                return;
            }

            OpenCalendarPopup();
        }

        private void OpenCalendarPopup()
        {
            var currentText = DateRangeText.Text;
            var yearText = YearText.Text;

            if (!string.IsNullOrEmpty(currentText) && currentText.Contains(" - "))
            {
                var parts = currentText.Split(new[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2 && int.TryParse(yearText, out int currentYear))
                {
                    if (DateTime.TryParseExact(parts[0].Trim() + "." + currentYear, "dd.MM.yyyy",
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime startDate) &&
                        DateTime.TryParseExact(parts[1].Trim() + "." + currentYear, "dd.MM.yyyy",
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime endDate))
                    {
                        SetDateInComboBoxes(startDate, endDate);
                    }
                }
            }

            UpdateCalendarDisplay(_currentCalendarDate);
            CalendarPopup.IsOpen = true;
        }

        private void ApplyDateRange_Click(object sender, RoutedEventArgs e)
        {
            var startDate = GetStartDateFromComboBoxes();
            var endDate = GetEndDateFromComboBoxes();

            if (startDate.HasValue && endDate.HasValue)
            {
                if (startDate > endDate)
                {
                    MessageBox.Show("Начальная дата не может быть позже конечной даты", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                UpdateDateDisplay(startDate.Value, endDate.Value);
                if (MainContentControl.Content is HistoryChangesView historyView)
                {
                    historyView.ApplyDateRangeFilter(startDate.Value, endDate.Value);
                }

                CalendarPopup.IsOpen = false;
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите обе даты", "Внимание",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void CancelDateRange_Click(object sender, RoutedEventArgs e)
        {
            CalendarPopup.IsOpen = false;
        }

        private string FormatDateForCalendar(DateTime date)
        {
            return $"{date:dd MM yy}";
        }

        private void UpdateDateDisplay(DateTime startDate, DateTime endDate)
        {
            DateRangeText.Text = $"{startDate:dd.MM} - {endDate:dd.MM}";

            if (startDate.Year == endDate.Year)
            {
                YearText.Text = startDate.Year.ToString();
            }
            else
            {
                YearText.Text = $"{startDate.Year}-{endDate.Year}";
            }

            OnDateRangeChanged(startDate, endDate);
        }

        private bool TryParseCalendarDate(string dateText, out DateTime result)
        {
            return DateTime.TryParseExact(dateText, "dd MM yy",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out result) ||
                   DateTime.TryParse(dateText, out result);
        }

        private void OnDateRangeChanged(DateTime startDate, DateTime endDate)
        {
            Console.WriteLine($"Период изменен: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}");

            if (startDate.Year == endDate.Year)
            {
                this.Title = $"Project BPI - Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy}";
            }
            else
            {
                this.Title = $"Project BPI - Период: {startDate:dd.MM.yyyy} - {endDate:dd.MM.yyyy} ({startDate.Year}-{endDate.Year})";
            }
        }

        Border currentActive = null;

        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj == null) yield break;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                if (child == null) continue;

                if (child is T t) yield return t;

                foreach (T childOfChild in FindVisualChildren<T>(child))
                    yield return childOfChild;
            }
        }

        private void ResetTextBlocksForeground(Border border)
        {
            if (border == null) return;

            foreach (TextBlock tb in FindVisualChildren<TextBlock>(border))
            {
                tb.ClearValue(TextBlock.ForegroundProperty);
            }
        }

        private void SetTextBlocksForeground(Border border, Brush brush)
        {
            if (border == null) return;

            foreach (TextBlock tb in FindVisualChildren<TextBlock>(border))
            {
                tb.Foreground = brush;
            }
        }

        private void ActivateMenuItem(Border border, bool showContent = true)
        {
            if (border == null)
            {
                return;
            }

            EnsureDynamicAncestorsVisible(border);
            ResetAllStyles();
            HighlightChain(border);
            currentActive = border;
            if (showContent)
            {
                ShowContentForMenuItem(border);
            }
            else if (dynamicTemplates.ContainsKey(border)
                || dynamicSections.ContainsKey(border)
                || dynamicSubSections.ContainsKey(border))
            {
                MainContentControl.Content = CreateDefaultContent();
            }
            UpdateArrows();
        }

        private void EnsureDynamicAncestorsVisible(Border border)
        {
            Border current = parents.ContainsKey(border)
                ? parents[border]
                : null;

            while (current != null)
            {
                if (dynamicMenus.TryGetValue(current, out var menu))
                {
                    menu.Visibility = Visibility.Visible;
                    if (dynamicIndicators.TryGetValue(current, out var indicator))
                    {
                        ApplyDynamicIndicatorState(indicator, true);
                    }
                }

                if (parents.ContainsKey(current))
                {
                    current = parents[current];
                }
                else
                {
                    break;
                }
            }
        }

        private void ResetAllStyles()
        {
            foreach (var border in FindVisualChildren<Border>(this))
            {
                string tag = border.Tag as string;

                if (tag == "main" || tag == "sub" || tag == "sub2")
                {
                    border.Style = (Style)FindResource("MenuItemStyle");
                    ResetTextBlocksForeground(border);
                }
            }
        }

        private void ApplyActiveStyle(Border b)
        {
            string tag = b.Tag as string;

            switch (tag)
            {
                case "main":
                    b.Style = (Style)FindResource("ActiveMainItemStyle");
                    break;
                case "sub":
                    b.Style = (Style)FindResource("ActiveSubItemStyle");
                    break;
                case "sub2":
                    b.Style = (Style)FindResource("ActiveSubItemLevel2Style");
                    break;
            }

            SetTextBlocksForeground(b, Brushes.White);
        }

        private void ShowContentForMenuItem(Border menuItem)
        {
            if (TryShowDynamicContent(menuItem))
            {
                UpdateHistoryCalendarVisibility(false);
                return;
            }

            UpdateHistoryCalendarVisibility(false);
            MainContentControl.Content = CreateDefaultContent();
        }

        private void HistoryButton_Click(object sender, RoutedEventArgs e)
        {
            var historyView = new HistoryChangesView(GetSharedTemplateDatabasePath());
            MainContentControl.Content = historyView;
            UpdateHistoryCalendarVisibility(true);

            var startDate = GetStartDateFromComboBoxes();
            var endDate = GetEndDateFromComboBoxes();
            if (startDate.HasValue && endDate.HasValue)
            {
                historyView.ApplyDateRangeFilter(startDate.Value, endDate.Value);
            }
        }
        private void HelpHyperlink_Click(object sender, RoutedEventArgs e)
        {
            UpdateHistoryCalendarVisibility(false);
            MainContentControl.Content = new Spravka_View();
        }
        private void HighlightChain(Border start)
        {
            Border current = start;
            while (current != null)
            {
                ApplyActiveStyle(current);
                if (parents.ContainsKey(current))
                    current = parents[current];
                else
                    break;
            }
        }

        private bool IsMenuItemActive(Border b)
        {
            return b.Style == (Style)FindResource("ActiveMainItemStyle")
                || b.Style == (Style)FindResource("ActiveSubItemStyle")
                || b.Style == (Style)FindResource("ActiveSubItemLevel2Style");
        }

        private void UpdateArrows()
        {
            foreach (var pair in dynamicIndicators)
            {
                bool isExpanded = dynamicMenus.TryGetValue(pair.Key, out var menu) && menu.Visibility == Visibility.Visible;
                ApplyIndicatorVisual(pair.Value, isExpanded, IsMenuItemActive(pair.Key));
            }
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SearchPlaceholder.Visibility = string.IsNullOrEmpty(SearchBox.Text) ? Visibility.Visible : Visibility.Collapsed;
            string query = SearchBox.Text?.Trim() ?? string.Empty;

            FilterMenuPanel(DynamicTemplatesPanel, query);
        }

        private void FilterMenuPanel(StackPanel panel, string query)
        {
            foreach (var child in panel.Children)
            {
                if (child is Border border)
                {
                    FilterMenuItem(border, query);

                    if (dynamicMenus.TryGetValue(border, out var subPanel))
                    {
                        FilterMenuPanel(subPanel, query);
                        bool hasVisibleChild = subPanel.Children.Cast<UIElement>().Any(c => c.Visibility == Visibility.Visible);
                        border.Visibility = (border.Visibility == Visibility.Visible || hasVisibleChild)
                            ? Visibility.Visible
                            : Visibility.Collapsed;
                    }
                }
                else if (child is StackPanel sp)
                {
                    FilterMenuPanel(sp, query);
                }
            }
        }

        private void FilterMenuItem(Border border, string query)
        {
            TextBlock tb = FindVisualChildren<TextBlock>(border).FirstOrDefault();

            if (tb != null)
            {
                bool match = string.IsNullOrEmpty(query)
                    || tb.Text?.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;
                border.Visibility = match ? Visibility.Visible : Visibility.Collapsed;
            }
        }
    }
}
