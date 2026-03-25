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
        private const string TemplateDatabaseFolderName = "TemplateDatabases";
        private const string SavedTemplatesFolderName = "SavedTemplates";
        private const string ArchivedReportsFolderName = "ArchivedReports";
        private const string SectionContentSubSectionTitle = "__section_content__";
        private const string StudyReportTitle = "Учебный отчет";
        private const string NirReportTitle = "Отчет по НИР";
        private const string NirPublishingTablePatternName = "nir_3_1_table";
        private const string StudyReportSuppressionKey = "builtin:study_report";
        private const string NirReportSuppressionKey = "builtin:nir_report";
        private const string HistoryActionCreate = "create";
        private const string HistoryActionEdit = "edit";
        private const string HistoryActionDelete = "delete";
        private sealed class StudyReportSubSectionSeed
        {
            public int Number { get; set; }
            public string Title { get; set; }
            public string TablePatternName { get; set; }
            public string[] Headers { get; set; }
        }

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

        private sealed class NirSubSectionSeed
        {
            public int Number { get; set; }
            public string Title { get; set; }
            public string TextPatternName { get; set; }
            public string TextContent { get; set; }
            public TableSeed[] Tables { get; set; }
            public NirSubSectionSeed[] Children { get; set; }
        }

        private sealed class NirSectionSeed
        {
            public int Number { get; set; }
            public string Title { get; set; }
            public string TextPatternName { get; set; }
            public string TextContent { get; set; }
            public TableSeed[] Tables { get; set; }
            public NirSubSectionSeed[] SubSections { get; set; }
        }

        private static readonly StudyReportSubSectionSeed[] StudyReportSubSections =
        {
            new StudyReportSubSectionSeed
            {
                Number = 1,
                Title = "Научно-издательская деятельность (факт)",
                TablePatternName = "study_14_1_table",
                Headers = new[]
                {
                    "№",
                    "Авторы",
                    "Количество",
                    "Тип публикации",
                    "Источник издания",
                    "Примечание"
                }
            },
            new StudyReportSubSectionSeed
            {
                Number = 2,
                Title = "Подача заявок на объекты интеллектуальной собственности (факт)",
                TablePatternName = "study_14_2_table",
                Headers = new[]
                {
                    "№",
                    "Тип ОИС",
                    "Авторы",
                    "Наименование",
                    "Дата подачи",
                    "№ патента или свидетельства",
                    "Примечание"
                }
            },
            new StudyReportSubSectionSeed
            {
                Number = 3,
                Title = "Проведения конференций, семинаров, совещаний (факт)",
                TablePatternName = "study_14_3_table",
                Headers = new[]
                {
                    "№",
                    "Название мероприятия",
                    "Уровень конференции",
                    "Дата и место проведения",
                    "Ответственный",
                    "Примечание"
                }
            }
        };

        private readonly Dictionary<Border, Border> parents = new Dictionary<Border, Border>();
        private readonly Dictionary<Border, DynamicTemplateEntry> dynamicTemplates = new Dictionary<Border, DynamicTemplateEntry>();
        private readonly Dictionary<Border, Section> dynamicSections = new Dictionary<Border, Section>();
        private readonly Dictionary<Border, SubSection> dynamicSubSections = new Dictionary<Border, SubSection>();
        private readonly Dictionary<Border, StackPanel> dynamicMenus = new Dictionary<Border, StackPanel>();
        private readonly Dictionary<Border, Image> dynamicIndicators = new Dictionary<Border, Image>();
        private int _applicationZoomPercent = 100;
        private DateTime _currentCalendarDate = DateTime.Today;
        private Border currentDynamicEditorBorder;

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

        private sealed class SectionEditorContext
        {
            public DynamicTemplateEntry Template { get; set; }
            public Section Section { get; set; }
            public TextBox TitleTextBox { get; set; }
            public TextBox ContentTextBox { get; set; }
            public SubSection ContentSubSection { get; set; }
            public List<TableEditorContext> TableEditors { get; } = new List<TableEditorContext>();
        }

        private sealed class SubSectionEditorContext
        {
            public DynamicTemplateEntry Template { get; set; }
            public SubSection SubSection { get; set; }
            public TextBox TitleTextBox { get; set; }
            public TextBox ContentTextBox { get; set; }
            public List<TableEditorContext> TableEditors { get; } = new List<TableEditorContext>();
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
            public bool FiltersEnabled { get; set; }
            public Dictionary<int, TableColumnFilterState> ColumnFilters { get; } = new Dictionary<int, TableColumnFilterState>();
        }

        private enum TableColumnSortMode
        {
            None,
            AlphabetAsc,
            AlphabetDesc,
            ValueAsc,
            ValueDesc
        }

        private const int MinimumApplicationZoomPercent = 50;
        private const int MaximumApplicationZoomPercent = 150;
        private const int ApplicationZoomStepPercent = 10;

        private sealed class TableColumnFilterState
        {
            public string SearchText { get; set; }
            public TableColumnSortMode SortMode { get; set; }

            public bool HasSettings =>
                !string.IsNullOrWhiteSpace(SearchText) || SortMode != TableColumnSortMode.None;
        }

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
                await EnsureNirReportInDatabaseAsync();
                await EnsureStudyReportInDatabaseAsync();
                await LoadSavedTemplatesAsync();
            };
        }

        private async void AddTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new TemplateNameDialog
            {
                Owner = this,
                Title = "Новый шаблон",
                Prompt = "Введите название нового шаблона",
                Label = "Название шаблона",
                TemplateName = $"Шаблон {DynamicTemplatesPanel.Children.Count + 1}"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            string templateTitle = dialog.TemplateName.Trim();
            if (dynamicTemplates.Values.Any(t =>
                string.Equals(t.DisplayTitle, templateTitle, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("Шаблон с таким названием уже существует.", "Шаблоны",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                var entry = await CreateTemplateDatabaseAsync(templateTitle);
                await LogHistoryAsync(
                    HistoryActionCreate,
                    "template",
                    BuildTemplateHistoryLocation(entry),
                    "Создан шаблон");
                ActivateMenuItem(entry.HeaderBorder, false);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось создать шаблон:{Environment.NewLine}{ex.Message}",
                    "Шаблоны", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private DynamicTemplateEntry AddTemplateToMenu(string templateTitle, Report report, string databasePath)
        {
            var templateContainer = new StackPanel();
            var templateMenu = new StackPanel { Visibility = Visibility.Collapsed };
            var templateHeader = CreateDynamicMenuBorder(templateTitle, "main", "MenuHeaderStyle", true, out var templateIndicator);

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
            dynamicMenus[templateHeader] = templateMenu;
            dynamicIndicators[templateHeader] = templateIndicator;
            parents[templateHeader] = null;
            entry.RegisteredBorders.Add(templateHeader);

            templateHeader.MouseLeftButtonDown += DynamicTemplateHeader_Click;

            templateContainer.Children.Add(templateHeader);
            templateContainer.Children.Add(templateMenu);

            foreach (var section in report.Sections.OrderBy(s => s.Number))
            {
                section.Report = report;

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
            if (string.Equals(templateTitle, NirReportTitle, StringComparison.Ordinal))
            {
                return 0;
            }

            if (string.Equals(templateTitle, StudyReportTitle, StringComparison.Ordinal))
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

        private async Task LoadSavedTemplatesAsync()
        {
            await LoadTemplatesFromDatabaseAsync(GetSharedTemplateDatabasePath());

            string folderPath = GetTemplateDatabaseFolderPath();
            foreach (string databasePath in Directory.GetFiles(folderPath, "*.db").OrderBy(System.IO.Path.GetFileName))
            {
                if (IsSharedTemplateDatabase(databasePath))
                {
                    continue;
                }

                await LoadTemplatesFromDatabaseAsync(databasePath);
            }
        }

        private async Task<DynamicTemplateEntry> CreateTemplateDatabaseAsync(string templateTitle)
        {
            string databasePath = GetSharedTemplateDatabasePath();
            var database = new DataBase(databasePath);
            database.InitializeDatabase(false);

            int patternId = await database.AddFilePattern(new PatternFile
            {
                Title = templateTitle,
                Year = DateTime.Today.Year,
                Path = databasePath
            });

            int reportId = await database.AddReport(new Report
            {
                Title = templateTitle,
                Year = DateTime.Today.Year,
                PattarnId = patternId
            });

            Report report = await database.GetFullReport(reportId);
            if (report == null)
            {
                throw new InvalidOperationException("Не удалось загрузить только что созданный шаблон из базы данных.");
            }

            return AddTemplateToMenu(templateTitle, report, databasePath);
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
                // Пропускаем поврежденную базу, чтобы остальные шаблоны загрузились.
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
            return templateEntry?.DisplayTitle?.Trim() ?? "Шаблон";
        }

        private string BuildSectionHistoryLocation(DynamicTemplateEntry templateEntry, Section section)
        {
            return $"{BuildTemplateHistoryLocation(templateEntry)} / {BuildSectionHistoryTitle(section)}";
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

            string sectionPart = BuildSectionHistoryTitle(subsection.Section);
            string subSectionPart = string.Join(" / ", parts);

            if (string.IsNullOrWhiteSpace(subSectionPart))
            {
                return $"{BuildTemplateHistoryLocation(templateEntry)} / {sectionPart}";
            }

            return $"{BuildTemplateHistoryLocation(templateEntry)} / {sectionPart} / {subSectionPart}";
        }

        private string BuildTableHistoryLocation(DynamicTemplateEntry templateEntry, SubSection subsection, Table table)
        {
            string tableTitle = string.IsNullOrWhiteSpace(table?.Title) ? "Таблица" : table.Title.Trim();
            string ownerLocation = subsection != null && IsSectionContentSubSection(subsection)
                ? BuildSectionHistoryLocation(templateEntry, subsection.Section)
                : BuildSubSectionHistoryLocation(templateEntry, subsection);

            return $"{ownerLocation} / {tableTitle}";
        }

        private string BuildHistoryDetails(params string[] fragments)
        {
            return string.Join(", ", fragments.Where(fragment => !string.IsNullOrWhiteSpace(fragment)));
        }

        private bool IsSharedTemplateDatabase(string databasePath)
        {
            return string.Equals(
                System.IO.Path.GetFullPath(databasePath),
                System.IO.Path.GetFullPath(GetSharedTemplateDatabasePath()),
                StringComparison.OrdinalIgnoreCase);
        }

        private bool IsTemplateLoaded(string databasePath, int reportId)
        {
            return dynamicTemplates.Values.Any(t =>
                string.Equals(t.DatabasePath, databasePath, StringComparison.OrdinalIgnoreCase) &&
                t.ReportId == reportId);
        }

        private string GetTemplateDatabaseFolderPath()
        {
            string folderPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, TemplateDatabaseFolderName);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            return folderPath;
        }

        private string GetSavedTemplatesFolderPath()
        {
            string folderPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, SavedTemplatesFolderName);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            return folderPath;
        }

        private string GetArchivedReportsFolderPath()
        {
            string folderPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ArchivedReportsFolderName);
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            return folderPath;
        }

        private string BuildSavedTemplateSnapshotPath(string templateTitle)
        {
            string safeTitle = Regex.Replace(templateTitle ?? "Шаблон", @"[^\w\dа-яА-Я_-]+", "_").Trim('_');
            if (string.IsNullOrWhiteSpace(safeTitle))
            {
                safeTitle = "template";
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return System.IO.Path.Combine(GetSavedTemplatesFolderPath(), $"{safeTitle}_{timestamp}.db");
        }

        private string BuildArchivedReportPath(string templateTitle)
        {
            string safeTitle = Regex.Replace(templateTitle ?? "Отчет", @"[^\w\dа-яА-Я_-]+", "_").Trim('_');
            if (string.IsNullOrWhiteSpace(safeTitle))
            {
                safeTitle = "report";
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return System.IO.Path.Combine(GetArchivedReportsFolderPath(), $"{safeTitle}_{timestamp}.db");
        }

        private string BuildRestoredTemplateDatabasePath(string templateTitle)
        {
            string safeTitle = Regex.Replace(templateTitle ?? "Шаблон", @"[^\w\dа-яА-Я_-]+", "_").Trim('_');
            if (string.IsNullOrWhiteSpace(safeTitle))
            {
                safeTitle = "template";
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return System.IO.Path.Combine(GetTemplateDatabaseFolderPath(), $"{safeTitle}_{timestamp}.db");
        }

        private async Task RestoreSavedTemplateSnapshotAsync(string snapshotPath)
        {
            if (string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
            {
                MessageBox.Show("Сохраненный шаблон не найден.", "Шаблоны",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var snapshotDatabase = new DataBase(snapshotPath);
                snapshotDatabase.InitializeDatabase(false);

                var snapshotReports = await snapshotDatabase.GetAllReports();
                var snapshotReportInfo = snapshotReports.FirstOrDefault();
                if (snapshotReportInfo == null)
                {
                    MessageBox.Show("Не удалось прочитать сохраненный шаблон.", "Шаблоны",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                Report snapshotReport = await snapshotDatabase.GetFullReport(snapshotReportInfo.Id);
                string restoredPath = BuildRestoredTemplateDatabasePath(snapshotReport?.Title);

                File.Copy(snapshotPath, restoredPath, false);

                var restoredDatabase = new DataBase(restoredPath);
                restoredDatabase.InitializeDatabase(false);

                var restoredReports = await restoredDatabase.GetAllReports();
                foreach (var restoredReportInfo in restoredReports)
                {
                    var restoredReport = await restoredDatabase.GetFullReport(restoredReportInfo.Id);
                    if (restoredReport?.PatternFile == null)
                    {
                        continue;
                    }

                    restoredReport.PatternFile.Path = restoredPath;
                    await restoredDatabase.UpdateFilePattern(restoredReport.PatternFile);
                }

                await LoadTemplatesFromDatabaseAsync(restoredPath);

                var loadedEntry = dynamicTemplates.Values.FirstOrDefault(entry =>
                    string.Equals(entry.DatabasePath, restoredPath, StringComparison.OrdinalIgnoreCase));
                if (loadedEntry != null)
                {
                    ActivateMenuItem(loadedEntry.HeaderBorder, false);
                }

                await LogHistoryAsync(
                    HistoryActionCreate,
                    "template_restore",
                    snapshotReport?.Title?.Trim() ?? "Шаблон",
                    "Выгружен сохраненный шаблон в меню навигации");

                MessageBox.Show(
                    $"Шаблон \"{snapshotReport?.Title}\" выгружен в меню навигации.",
                    "Шаблоны",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось выгрузить шаблон:{Environment.NewLine}{ex.Message}",
                    "Шаблоны",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async Task DeleteSavedTemplateSnapshotAsync(string snapshotPath)
        {
            if (string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
            {
                return;
            }

            try
            {
                string templateTitle = "Шаблон";
                var snapshotDatabase = new DataBase(snapshotPath);
                snapshotDatabase.InitializeDatabase(false);

                var reports = await snapshotDatabase.GetAllReports();
                var reportInfo = reports.FirstOrDefault();
                if (reportInfo != null)
                {
                    templateTitle = reportInfo.Title;
                }

                File.Delete(snapshotPath);

                await LogHistoryAsync(
                    HistoryActionDelete,
                    "template_snapshot",
                    templateTitle,
                    "Удален сохраненный шаблон из вкладки «Шаблоны»");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось удалить сохраненный шаблон:{Environment.NewLine}{ex.Message}",
                    "Шаблоны",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async Task<Report> CloneReportToSnapshotAsync(string sourceDatabasePath, int reportId, string snapshotPath)
        {
            var sourceDatabase = new DataBase(sourceDatabasePath);
            sourceDatabase.InitializeDatabase(false);

            Report sourceReport = await sourceDatabase.GetFullReport(reportId);
            if (sourceReport == null)
            {
                return null;
            }

            var snapshotDatabase = new DataBase(snapshotPath);
            snapshotDatabase.InitializeDatabase(false);

            int snapshotPatternId = await snapshotDatabase.AddFilePattern(new PatternFile
            {
                Title = sourceReport.Title,
                Year = sourceReport.Year,
                Path = snapshotPath
            });

            int snapshotReportId = await snapshotDatabase.AddReport(new Report
            {
                Title = sourceReport.Title,
                Year = sourceReport.Year,
                PattarnId = snapshotPatternId
            });

            foreach (var section in sourceReport.Sections.OrderBy(section => section.Number).ThenBy(section => section.Id))
            {
                int snapshotSectionId = await snapshotDatabase.AddSection(new Section
                {
                    ReportId = snapshotReportId,
                    Number = section.Number,
                    Title = section.Title
                });

                await CloneSubSectionsAsync(
                    snapshotDatabase,
                    snapshotSectionId,
                    section.SubSections?.OrderBy(item => item.Number).ThenBy(item => item.Id) ?? Enumerable.Empty<SubSection>(),
                    null);
            }

            return sourceReport;
        }

        private async Task SaveTemplateSnapshotAsync(Border ownerBorder)
        {
            if (ownerBorder == null || !dynamicTemplates.TryGetValue(ownerBorder, out var templateEntry))
            {
                return;
            }

            string snapshotPath = BuildSavedTemplateSnapshotPath(templateEntry.DisplayTitle);

            try
            {
                Report sourceReport = await CloneReportToSnapshotAsync(
                    templateEntry.DatabasePath,
                    templateEntry.ReportId,
                    snapshotPath);
                if (sourceReport == null)
                {
                    MessageBox.Show("Не удалось загрузить шаблон для сохранения.", "Шаблоны",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                await LogHistoryAsync(
                    HistoryActionCreate,
                    "template_snapshot",
                    BuildTemplateHistoryLocation(templateEntry),
                    "Сохранен снимок шаблона во вкладку «Шаблоны»");

                if (MainContentControl.Content is TemplatesPage templatesPage)
                {
                    await templatesPage.ReloadAsync();
                }

                MessageBox.Show($"Шаблон \"{templateEntry.DisplayTitle}\" сохранен во вкладку «Шаблоны».",
                    "Шаблоны", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось сохранить шаблон:{Environment.NewLine}{ex.Message}",
                    "Шаблоны", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

        private async void GenerateFinalReportButton_Click(object sender, RoutedEventArgs e)
        {
            var templateEntry = GetCurrentTemplateEntry();
            if (templateEntry == null)
            {
                MessageBox.Show("Выберите отчет, который нужно перенести в архив.", "Архив",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string archivePath = BuildArchivedReportPath(templateEntry.DisplayTitle);

            try
            {
                Report archivedReport = await CloneReportToSnapshotAsync(
                    templateEntry.DatabasePath,
                    templateEntry.ReportId,
                    archivePath);
                if (archivedReport == null)
                {
                    MessageBox.Show("Не удалось сформировать итоговый отчет.", "Архив",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                await RemoveTemplateSourceAsync(templateEntry, archivedReport);
                RemoveTemplateFromMenu(templateEntry);
                MainContentControl.Content = CreateDefaultContent();
                currentDynamicEditorBorder = null;

                await LogHistoryAsync(
                    HistoryActionCreate,
                    "archive_report",
                    archivedReport.Title,
                    "Отчет перенесен в архив");

                MessageBox.Show(
                    $"Отчет \"{archivedReport.Title}\" перенесен в архив.",
                    "Архив",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось перенести отчет в архив:{Environment.NewLine}{ex.Message}",
                    "Архив",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async Task RemoveTemplateSourceAsync(DynamicTemplateEntry templateEntry, Report sourceReport)
        {
            if (templateEntry == null || sourceReport == null)
            {
                return;
            }

            var database = new DataBase(templateEntry.DatabasePath);
            database.InitializeDatabase(false);

            if (IsSharedTemplateDatabase(templateEntry.DatabasePath))
            {
                if (string.Equals(sourceReport.Title, NirReportTitle, StringComparison.Ordinal))
                {
                    await database.AddSuppressedBuiltInReport(NirReportSuppressionKey);
                }
                else if (string.Equals(sourceReport.Title, StudyReportTitle, StringComparison.Ordinal))
                {
                    await database.AddSuppressedBuiltInReport(StudyReportSuppressionKey);
                }

                await database.DeleteFilePattern(sourceReport.PattarnId);
                return;
            }

            var reports = await database.GetAllReports();
            if (reports.Count <= 1)
            {
                if (File.Exists(templateEntry.DatabasePath))
                {
                    File.Delete(templateEntry.DatabasePath);
                }

                return;
            }

            await database.DeleteFilePattern(sourceReport.PattarnId);
        }

        private async Task DownloadArchivedReportAsync(string archivePath)
        {
            if (string.IsNullOrWhiteSpace(archivePath) || !File.Exists(archivePath))
            {
                MessageBox.Show("Архивный отчет не найден.", "Архив",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                var database = new DataBase(archivePath);
                database.InitializeDatabase(false);

                var reports = await database.GetAllReports();
                var reportInfo = reports.FirstOrDefault();
                if (reportInfo == null)
                {
                    MessageBox.Show("Не удалось загрузить архивный отчет.", "Архив",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var report = await database.GetFullReport(reportInfo.Id);
                if (report == null)
                {
                    MessageBox.Show("Не удалось загрузить архивный отчет.", "Архив",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var saveDialog = new SaveFileDialog
                {
                    Title = "Сохранить итоговый отчет",
                    Filter = "Документ Word (*.docx)|*.docx",
                    FileName = $"{SanitizeFileName(report.Title)}.docx",
                    DefaultExt = ".docx",
                    AddExtension = true
                };

                if (saveDialog.ShowDialog() != true)
                {
                    return;
                }

                DocxExportService.ExportReport(report, saveDialog.FileName);

                await LogHistoryAsync(
                    HistoryActionCreate,
                    "archive_download",
                    report.Title,
                    "Архивный отчет выгружен в формате DOCX");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось скачать архивный отчет:{Environment.NewLine}{ex.Message}",
                    "Архив",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async Task DeleteArchivedReportAsync(string archivePath)
        {
            if (string.IsNullOrWhiteSpace(archivePath) || !File.Exists(archivePath))
            {
                return;
            }

            try
            {
                string title = "Отчет";
                var database = new DataBase(archivePath);
                database.InitializeDatabase(false);
                var reports = await database.GetAllReports();
                if (reports.Any())
                {
                    title = reports.First().Title;
                }

                File.Delete(archivePath);

                await LogHistoryAsync(
                    HistoryActionDelete,
                    "archive_report",
                    title,
                    "Архивный отчет удален");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Не удалось удалить архивный отчет:{Environment.NewLine}{ex.Message}",
                    "Архив",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private string SanitizeFileName(string name)
        {
            string sanitized = Regex.Replace(name ?? "report", $"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()))}]", "_");
            return string.IsNullOrWhiteSpace(sanitized) ? "report" : sanitized.Trim();
        }

        private async Task CloneSubSectionsAsync(
            DataBase targetDatabase,
            int targetSectionId,
            IEnumerable<SubSection> sourceSubSections,
            int? targetParentSubSectionId)
        {
            foreach (var sourceSubSection in sourceSubSections ?? Enumerable.Empty<SubSection>())
            {
                int snapshotSubSectionId = await targetDatabase.AddSubsection(new SubSection
                {
                    SectionId = targetSectionId,
                    ParentSubsectionId = targetParentSubSectionId,
                    Number = sourceSubSection.Number,
                    Title = sourceSubSection.Title
                });

                foreach (var text in sourceSubSection.Texts?.OrderBy(item => item.Id) ?? Enumerable.Empty<Text>())
                {
                    await targetDatabase.AddText(new Text
                    {
                        SubsectionId = snapshotSubSectionId,
                        Content = text.Content,
                        PatternName = text.PatternName
                    });
                }

                foreach (var table in sourceSubSection.Tables?.OrderBy(item => item.Id) ?? Enumerable.Empty<Table>())
                {
                    int snapshotTableId = await targetDatabase.AddTable(new Table
                    {
                        Title = table.Title,
                        SubsectionId = snapshotSubSectionId,
                        PatternName = table.PatternName
                    });

                    foreach (var cell in table.TableItems?.OrderBy(item => item.IsHeader ? 0 : 1).ThenBy(item => item.Row).ThenBy(item => item.Column) ?? Enumerable.Empty<TableItem>())
                    {
                        await targetDatabase.AddTableItem(new TableItem
                        {
                            TableId = snapshotTableId,
                            Row = cell.Row,
                            Column = cell.Column,
                            Header = cell.Header,
                            ColSpan = cell.ColSpan,
                            RowSpan = cell.RowSpan,
                            IsHeader = cell.IsHeader
                        });
                    }
                }

                await CloneSubSectionsAsync(
                    targetDatabase,
                    targetSectionId,
                    sourceSubSection.SubSections?.OrderBy(item => item.Number).ThenBy(item => item.Id) ?? Enumerable.Empty<SubSection>(),
                    snapshotSubSectionId);
            }
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
            int? subsectionId = null,
            bool editMode = true)
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
                ActivateMenuItem(selectedBorder, true, editMode);
            }
        }

        private async Task EnsureStudyReportInDatabaseAsync()
        {
            var database = new DataBase(GetSharedTemplateDatabasePath());
            database.InitializeDatabase(false);

            if (await database.HasSuppressedBuiltInReport(StudyReportSuppressionKey))
            {
                return;
            }

            var reports = await database.GetAllReports();
            var existingStudyReport = reports.FirstOrDefault(report =>
                string.Equals(report.Title, StudyReportTitle, StringComparison.Ordinal));

            int reportId;
            if (existingStudyReport == null)
            {
                int patternId = await database.AddFilePattern(new PatternFile
                {
                    Title = StudyReportTitle,
                    Year = DateTime.Today.Year,
                    Path = database.DatabasePath
                });

                reportId = await database.AddReport(new Report
                {
                    Title = StudyReportTitle,
                    Year = DateTime.Today.Year,
                    PattarnId = patternId
                });
            }
            else
            {
                reportId = existingStudyReport.Id;
            }

            var studyReport = await database.GetFullReport(reportId);
            if (studyReport == null)
            {
                return;
            }

            var section14 = studyReport.Sections.FirstOrDefault(section => section.Number == 14);
            if (section14 == null)
            {
                int sectionId = await database.AddSection(new Section
                {
                    ReportId = reportId,
                    Number = 14,
                    Title = GetDefaultSectionTitle(14)
                });

                studyReport = await database.GetFullReport(reportId);
                section14 = studyReport?.Sections.FirstOrDefault(section => section.Id == sectionId);
            }

            if (section14 == null)
            {
                return;
            }

            bool addedSubSections = false;
            foreach (var seed in StudyReportSubSections)
            {
                var subsection = section14.SubSections.FirstOrDefault(item =>
                    !item.ParentSubsectionId.HasValue &&
                    (item.Number == seed.Number || string.Equals(item.Title, seed.Title, StringComparison.Ordinal)));

                if (subsection != null)
                {
                    continue;
                }

                await database.AddSubsection(new SubSection
                {
                    SectionId = section14.Id,
                    ParentSubsectionId = null,
                    Number = seed.Number,
                    Title = seed.Title
                });
                addedSubSections = true;
            }

            if (addedSubSections)
            {
                studyReport = await database.GetFullReport(reportId);
                section14 = studyReport?.Sections.FirstOrDefault(section => section.Number == 14);
            }

            if (section14 == null)
            {
                return;
            }

            foreach (var seed in StudyReportSubSections)
            {
                var subsection = section14.SubSections.FirstOrDefault(item =>
                    !item.ParentSubsectionId.HasValue &&
                    string.Equals(item.Title, seed.Title, StringComparison.Ordinal));

                if (subsection == null)
                {
                    continue;
                }

                var existingTable = subsection.Tables.FirstOrDefault(table =>
                    string.Equals(table.PatternName, seed.TablePatternName, StringComparison.Ordinal));

                if (existingTable != null)
                {
                    continue;
                }

                int tableId = await database.AddTable(new Table
                {
                    Title = seed.Title,
                    SubsectionId = subsection.Id,
                    PatternName = seed.TablePatternName
                });

                for (int column = 0; column < seed.Headers.Length; column++)
                {
                    await database.AddTableItem(new TableItem
                    {
                        TableId = tableId,
                        Row = 1,
                        Column = column + 1,
                        Header = seed.Headers[column],
                        ColSpan = 1,
                        RowSpan = 1,
                        IsHeader = true
                    });
                }
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

            var editButton = CreateDynamicEditButton(border);
            var deleteButton = CreateDynamicDeleteButton(border);
            var saveButton = string.Equals(tag, "main", StringComparison.Ordinal)
                ? CreateDynamicSaveButton(border)
                : null;
            var actionsPanel = saveButton != null
                ? CreateDynamicActionsPanel(saveButton, editButton, deleteButton)
                : CreateDynamicActionsPanel(editButton, deleteButton);

            border.MouseEnter += (sender, args) =>
            {
                if (saveButton != null)
                {
                    saveButton.Visibility = Visibility.Visible;
                }
                editButton.Visibility = Visibility.Visible;
                deleteButton.Visibility = Visibility.Visible;
            };
            border.MouseLeave += (sender, args) =>
            {
                if (saveButton != null)
                {
                    saveButton.Visibility = Visibility.Hidden;
                }
                editButton.Visibility = Visibility.Hidden;
                deleteButton.Visibility = Visibility.Hidden;
            };

            if (!expandable)
            {
                var leafGrid = new Grid();
                leafGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                leafGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                Grid.SetColumn(title, 0);
                Grid.SetColumn(actionsPanel, 1);

                leafGrid.Children.Add(title);
                leafGrid.Children.Add(actionsPanel);
                border.Child = leafGrid;
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

            Grid.SetColumn(titlePanel, 0);
            Grid.SetColumn(actionsPanel, 1);

            grid.Children.Add(titlePanel);
            grid.Children.Add(actionsPanel);
            border.Child = grid;

            return border;
        }

        private StackPanel CreateDynamicActionsPanel(params Button[] buttons)
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Center
            };

            foreach (var button in buttons)
            {
                panel.Children.Add(button);
            }

            return panel;
        }

        private Button CreateDynamicEditButton(Border ownerBorder)
        {
            var button = CreateDynamicIconButton("Редактировать");
            button.Content = CreateDynamicEditIcon();
            button.PreviewMouseLeftButtonDown += (sender, args) => args.Handled = true;
            button.PreviewMouseLeftButtonUp += (sender, args) =>
            {
                args.Handled = true;
                ToggleDynamicEditor(ownerBorder);
            };

            return button;
        }

        private Button CreateDynamicDeleteButton(Border ownerBorder)
        {
            var button = CreateDynamicIconButton("Удалить");
            button.Content = CreateDynamicDeleteIcon();
            button.PreviewMouseLeftButtonDown += (sender, args) => args.Handled = true;
            button.PreviewMouseLeftButtonUp += async (sender, args) =>
            {
                args.Handled = true;
                await DeleteDynamicItemAsync(ownerBorder);
            };

            return button;
        }

        private Button CreateDynamicIconButton(string toolTip)
        {
            return new Button
            {
                Visibility = Visibility.Hidden,
                Width = 20,
                Height = 20,
                Margin = new Thickness(2, 0, 2, 0),
                Padding = new Thickness(0),
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Focusable = false,
                ToolTip = toolTip,
                Template = CreateDynamicEditButtonTemplate()
            };
        }

        private void ToggleDynamicEditor(Border ownerBorder)
        {
            if (ownerBorder == null)
            {
                return;
            }

            if (currentDynamicEditorBorder == ownerBorder)
            {
                MainContentControl.Content = CreateDefaultContent();
                currentDynamicEditorBorder = null;
                return;
            }

            ActivateMenuItem(ownerBorder, true, true);
        }

        private ControlTemplate CreateDynamicEditButtonTemplate()
        {
            var borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.Name = "EditButtonBorder";
            borderFactory.SetValue(Border.BackgroundProperty, new SolidColorBrush((Color)ColorConverter.ConvertFromString("#EEF4FF")));
            borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(12));
            borderFactory.SetValue(Border.PaddingProperty, new Thickness(4));

            var presenterFactory = new FrameworkElementFactory(typeof(ContentPresenter));
            presenterFactory.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            presenterFactory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            borderFactory.AppendChild(presenterFactory);

            var template = new ControlTemplate(typeof(Button))
            {
                VisualTree = borderFactory
            };

            var hoverTrigger = new Trigger
            {
                Property = UIElement.IsMouseOverProperty,
                Value = true
            };
            hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty,
                new SolidColorBrush((Color)ColorConverter.ConvertFromString("#DCE8FF")), "EditButtonBorder"));

            var pressedTrigger = new Trigger
            {
                Property = Button.IsPressedProperty,
                Value = true
            };
            pressedTrigger.Setters.Add(new Setter(Border.BackgroundProperty,
                new SolidColorBrush((Color)ColorConverter.ConvertFromString("#C8DBFF")), "EditButtonBorder"));

            template.Triggers.Add(hoverTrigger);
            template.Triggers.Add(pressedTrigger);
            return template;
        }

        private Viewbox CreateDynamicEditIcon()
        {
            var canvas = new Canvas
            {
                Width = 24,
                Height = 24
            };

            Brush iconBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2E4491"));

            canvas.Children.Add(new System.Windows.Shapes.Path
            {
                Data = Geometry.Parse("M21.707,5.565,18.435,2.293a1,1,0,0,0-1.414,0L3.93,15.384a.991.991,0,0,0-.242.39l-1.636,4.91A1,1,0,0,0,3,22a.987.987,0,0,0,.316-.052l4.91-1.636a.991.991,0,0,0,.39-.242L21.707,6.979A1,1,0,0,0,21.707,5.565ZM7.369,18.489l-2.788.93.93-2.788,8.943-8.944,1.859,1.859ZM17.728,8.132l-1.86-1.86,1.86-1.858,1.858,1.858Z"),
                Fill = iconBrush,
                Stretch = Stretch.Fill
            });

            return new Viewbox
            {
                Width = 12,
                Height = 12,
                Stretch = Stretch.Uniform,
                Child = canvas
            };
        }

        private Viewbox CreateDynamicDeleteIcon()
        {
            var canvas = new Canvas
            {
                Width = 24,
                Height = 24
            };

            Brush iconBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#8A1F1F"));

            canvas.Children.Add(new System.Windows.Shapes.Path
            {
                Data = Geometry.Parse("M7,18V14a1,1,0,0,1,2,0v4a1,1,0,0,1-2,0Zm5,1a1,1,0,0,0,1-1V14a1,1,0,0,0-2,0v4A1,1,0,0,0,12,19Zm4,0a1,1,0,0,0,1-1V14a1,1,0,0,0-2,0v4A1,1,0,0,0,16,19ZM23,6v4a1,1,0,0,1-1,1H21V22a1,1,0,0,1-1,1H4a1,1,0,0,1-1-1V11H2a1,1,0,0,1-1-1V6A1,1,0,0,1,2,5H7V2A1,1,0,0,1,8,1h8a1,1,0,0,1,1,1V5h5A1,1,0,0,1,23,6ZM9,5h6V3H9Zm10,6H5V21H19Zm2-4H3V9H21Z"),
                Fill = iconBrush,
                Stretch = Stretch.Fill
            });

            return new Viewbox
            {
                Width = 12,
                Height = 12,
                Stretch = Stretch.Uniform,
                Child = canvas
            };
        }

        private Viewbox CreateDynamicSaveIcon()
        {
            var canvas = new Canvas
            {
                Width = 24,
                Height = 24
            };

            Brush iconBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#758CA3"));

            canvas.Children.Add(new System.Windows.Shapes.Path
            {
                Data = Geometry.Parse("M5 0V7C5 7.55228 5.44772 8 6 8H16C16.5523 8 17 7.55228 17 7V0H18L22 5V21C22 22.6569 20.6569 24 19 24H3C1.34315 24 0 22.6569 0 21V3C0 1.34315 1.34315 0 3 0H5zM7 0H15V6H7V0zM6.2 15H15.8C16.4627 15 17 14.5523 17 14C17 13.4477 16.4627 13 15.8 13H6.2C5.53726 13 5 13.4477 5 14C5 14.5523 5.53726 15 6.2 15zM6.2 19H15.8C16.4627 19 17 18.5523 17 18C17 17.4477 16.4627 17 15.8 17H6.2C5.53726 17 5 17.4477 5 18C5 18.5523 5.53726 19 6.2 19z"),
                Fill = iconBrush,
                Stretch = Stretch.Fill
            });

            return new Viewbox
            {
                Width = 12,
                Height = 12,
                Stretch = Stretch.Uniform,
                Child = canvas
            };
        }

        private async Task DeleteDynamicItemAsync(Border ownerBorder)
        {
            if (ownerBorder == null)
            {
                return;
            }

            if (dynamicTemplates.TryGetValue(ownerBorder, out var templateEntry))
            {
                await DeleteTemplateAsync(templateEntry);
                return;
            }

            if (dynamicSections.TryGetValue(ownerBorder, out var section))
            {
                var template = FindTemplateEntryForSection(section);
                if (template != null)
                {
                    await DeleteSectionAsync(template, section);
                }

                return;
            }

            if (dynamicSubSections.TryGetValue(ownerBorder, out var subsection))
            {
                var template = FindTemplateEntryForSubSection(subsection);
                if (template != null)
                {
                    await DeleteSubSectionAsync(template, subsection);
                }
            }
        }

        private void DynamicTemplateHeader_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border)
            {
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

                ActivateMenuItem(border, true, false);
            }
        }

        private void DynamicLeaf_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border)
            {
                ActivateMenuItem(border, true, false);
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

                ActivateMenuItem(border, true, false);
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

        private bool TryShowDynamicContent(Border menuItem, bool editMode)
        {
            if (dynamicTemplates.TryGetValue(menuItem, out var templateEntry))
            {
                if (!editMode)
                {
                    return false;
                }

                MainContentControl.Content = editMode
                    ? CreateTemplateContent(templateEntry)
                    : CreateTemplatePreviewContent(templateEntry);
                currentDynamicEditorBorder = editMode ? menuItem : null;
                return true;
            }

            if (dynamicSections.TryGetValue(menuItem, out var section))
            {
                if (!editMode && !HasSectionDisplayContent(section))
                {
                    MainContentControl.Content = CreateDefaultContent();
                    currentDynamicEditorBorder = null;
                    return true;
                }

                MainContentControl.Content = editMode
                    ? CreateSectionContent(section)
                    : CreateSectionPreviewContent(section);
                currentDynamicEditorBorder = editMode ? menuItem : null;
                return true;
            }

            if (dynamicSubSections.TryGetValue(menuItem, out var subsection))
            {
                if (!editMode && !HasSubSectionDisplayContent(subsection))
                {
                    MainContentControl.Content = CreateDefaultContent();
                    currentDynamicEditorBorder = null;
                    return true;
                }

                MainContentControl.Content = editMode
                    ? CreateSubSectionContent(subsection)
                    : CreateSubSectionPreviewContent(subsection);
                currentDynamicEditorBorder = editMode ? menuItem : null;
                return true;
            }

            return false;
        }

        private UIElement CreateTemplatePreviewContent(DynamicTemplateEntry templateEntry)
        {
            var stack = CreateContentStack(templateEntry.DisplayTitle, GetTemplateStorageDescription(templateEntry));

            string sectionSummary = templateEntry.Report.Sections.Any()
                ? string.Join(Environment.NewLine, templateEntry.Report.Sections
                    .OrderBy(section => section.Number)
                    .Select(BuildSectionTitle))
                : "Разделы еще не созданы.";

            stack.Children.Add(CreateTextCard("Разделы шаблона", sectionSummary));

            if (templateEntry.Report.Sections.Any())
            {
                foreach (var section in templateEntry.Report.Sections.OrderBy(section => section.Number))
                {
                    var visibleSubSections = GetVisibleSubSections(section).ToList();
                    string sectionText = GetSectionDisplayContent(section);
                    string subSectionSummary = visibleSubSections.Any()
                        ? string.Join(Environment.NewLine, BuildSubSectionOutlineLines(section, visibleSubSections))
                        : "Подразделы отсутствуют.";

                    stack.Children.Add(CreateTextCard(
                        BuildSectionTitle(section),
                        $"{sectionText}{Environment.NewLine}{Environment.NewLine}Подразделы:{Environment.NewLine}{subSectionSummary}"));
                }
            }

            return new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = stack
            };
        }

        private UIElement CreateTemplateContent(DynamicTemplateEntry templateEntry)
        {
            var stack = CreateContentStack(templateEntry.DisplayTitle,
                GetTemplateStorageDescription(templateEntry));

            var templateButtons = CreateButtonRow();

            var addSectionButton = CreateActionButton("Добавить раздел");
            addSectionButton.Tag = templateEntry;
            addSectionButton.Click += AddSectionButton_Click;
            templateButtons.Children.Add(addSectionButton);

            var renameTemplateButton = CreateSecondaryButton("Переименовать шаблон");
            renameTemplateButton.Tag = templateEntry;
            renameTemplateButton.Click += RenameTemplateButton_Click;
            templateButtons.Children.Add(renameTemplateButton);

            var deleteTemplateButton = CreateDangerButton("Удалить шаблон");
            deleteTemplateButton.Tag = templateEntry;
            deleteTemplateButton.Click += DeleteTemplateButton_Click;
            templateButtons.Children.Add(deleteTemplateButton);

            stack.Children.Add(templateButtons);

            if (templateEntry.Report.Sections.Any())
            {
                foreach (var section in templateEntry.Report.Sections.OrderBy(s => s.Number))
                {
                    var visibleSubSections = GetVisibleSubSections(section).ToList();
                    var sectionLines = visibleSubSections.Any()
                        ? BuildSubSectionOutlineLines(section, visibleSubSections)
                        : new[] { "Подразделы отсутствуют" };

                    stack.Children.Add(CreateTextCard(BuildSectionTitle(section), string.Join(Environment.NewLine, sectionLines)));
                }
            }
            else
            {
                stack.Children.Add(CreateTextCard("Разделы", "Разделы еще не созданы. Добавьте первый раздел кнопкой выше."));
            }

            return new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = stack
            };
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

        private UIElement CreateSectionContent(Section section)
        {
            var templateEntry = FindTemplateEntryForSection(section);
            var contentSubSection = GetSectionContentSubSection(section);
            var editorContext = new SectionEditorContext
            {
                Template = templateEntry,
                Section = section,
                ContentSubSection = contentSubSection
            };
            var stack = CreateContentStack(GetSectionContentTitle(section),
                "Здесь можно изменить название раздела, добавить текст, таблицы и новые подразделы.");

            var titleBox = CreateEditorTextBox(section.Title, false);
            editorContext.TitleTextBox = titleBox;
            stack.Children.Add(CreateContentCard("Название раздела", titleBox));

            string content = contentSubSection?.Texts != null && contentSubSection.Texts.Any()
                ? contentSubSection.Texts.First().Content
                : string.Empty;
            var contentBox = CreateEditorTextBox(content, true);
            editorContext.ContentTextBox = contentBox;
            stack.Children.Add(CreateContentCard("Текст раздела", contentBox));

            if (contentSubSection?.Tables != null && contentSubSection.Tables.Any())
            {
                foreach (var table in contentSubSection.Tables)
                {
                    stack.Children.Add(CreateTableEditorCard(templateEntry, contentSubSection, table, editorContext.TableEditors));
                }
            }

            var buttons = CreateButtonRow();

            var saveButton = CreateActionButton("Сохранить");
            saveButton.Tag = editorContext;
            saveButton.Click += SaveSectionButton_Click;
            buttons.Children.Add(saveButton);

            var addTableButton = CreateSecondaryButton("Добавить таблицу");
            addTableButton.Tag = editorContext;
            addTableButton.Click += AddTableButton_Click;
            buttons.Children.Add(addTableButton);

            var addSubSectionButton = CreateSecondaryButton("Добавить подраздел");
            addSubSectionButton.Tag = editorContext;
            addSubSectionButton.Click += AddSubSectionButton_Click;
            buttons.Children.Add(addSubSectionButton);

            stack.Children.Add(buttons);

            return new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = stack
            };
        }

        private UIElement CreateSubSectionContent(SubSection subsection)
        {
            var templateEntry = FindTemplateEntryForSubSection(subsection);
            string title = BuildSubSectionTitle(subsection.Section, subsection);
            string storageDescription = templateEntry != null && IsSharedTemplateDatabase(templateEntry.DatabasePath)
                ? "Изменения сохраняются в общую базу данных Kurs.db."
                : "Изменения сохраняются в отдельную базу данных шаблона.";
            bool canAddNestedSubSection = !subsection.ParentSubsectionId.HasValue;
            var editorContext = new SubSectionEditorContext
            {
                Template = templateEntry,
                SubSection = subsection
            };
            var stack = CreateContentStack(title,
                storageDescription);

            var titleBox = CreateEditorTextBox(subsection.Title, false);
            editorContext.TitleTextBox = titleBox;
            stack.Children.Add(CreateContentCard("Название подраздела", titleBox));

            string content = subsection.Texts != null && subsection.Texts.Any()
                ? subsection.Texts.First().Content
                : string.Empty;

            var contentBox = CreateEditorTextBox(content, true);
            editorContext.ContentTextBox = contentBox;
            stack.Children.Add(CreateContentCard("Текст подраздела", contentBox));

            if (subsection.Tables != null && subsection.Tables.Any())
            {
                foreach (var table in subsection.Tables)
                {
                    stack.Children.Add(CreateTableEditorCard(templateEntry, subsection, table, editorContext.TableEditors));
                }
            }

            var saveButton = CreateActionButton("Сохранить");
            saveButton.Tag = editorContext;
            saveButton.Click += SaveSubSectionButton_Click;
            var subsectionButtons = CreateButtonRow();
            subsectionButtons.Children.Add(saveButton);

            var addTableButton = CreateSecondaryButton("Добавить таблицу");
            addTableButton.Tag = editorContext;
            addTableButton.Click += AddTableButton_Click;
            subsectionButtons.Children.Add(addTableButton);

            if (canAddNestedSubSection)
            {
                var addNestedSubSectionButton = CreateSecondaryButton("Добавить вложенный подраздел");
                addNestedSubSectionButton.Tag = editorContext;
                addNestedSubSectionButton.Click += AddSubSectionButton_Click;
                subsectionButtons.Children.Add(addNestedSubSectionButton);
            }

            stack.Children.Add(subsectionButtons);

            return new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = stack
            };
        }

        private UIElement CreateSubSectionPreviewContent(SubSection subsection)
        {
            var templateEntry = FindTemplateEntryForSubSection(subsection);
            string sectionTitle = subsection.Section != null
                ? GetSectionContentTitle(subsection.Section)
                : string.Empty;
            string subSectionTitle = BuildSubSectionTitle(subsection.Section, subsection);
            string content = subsection.Texts != null && subsection.Texts.Any()
                ? subsection.Texts.First().Content
                : string.Empty;

            var stack = new StackPanel
            {
                Margin = new Thickness(20)
            };

            stack.Children.Add(CreateSubSectionPreviewHeader(sectionTitle, subSectionTitle));
            if (!string.IsNullOrWhiteSpace(content))
            {
                stack.Children.Add(CreateSubSectionPreviewBody(content));
            }

            if (subsection.Tables != null && subsection.Tables.Any())
            {
                foreach (var table in subsection.Tables.OrderBy(t => t.Id))
                {
                    stack.Children.Add(CreateSubSectionPreviewTableCard(templateEntry, subsection, table));
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

            string title = section.Title.Trim();
            if (section.Number <= 0)
            {
                return title;
            }

            return string.Equals(title, GetDefaultSectionTitle(section.Number), StringComparison.Ordinal)
                ? string.Empty
                : $"{section.Number} {title}";
        }

        private Border CreateSubSectionPreviewHeader(string sectionTitle, string subSectionTitle)
        {
            var panel = new StackPanel
            {
                HorizontalAlignment = HorizontalAlignment.Center
            };

            if (!string.IsNullOrWhiteSpace(sectionTitle))
            {
                panel.Children.Add(new TextBlock
                {
                    Text = sectionTitle,
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.White,
                    TextAlignment = TextAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 10),
                    TextWrapping = TextWrapping.Wrap
                });
            }

            panel.Children.Add(new TextBlock
            {
                Text = subSectionTitle,
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                TextAlignment = TextAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 5)
            });

            return new Border
            {
                CornerRadius = new CornerRadius(35),
                Margin = new Thickness(10),
                Padding = new Thickness(15),
                Height = 80,
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

        private Border CreateSubSectionPreviewTableCard(Table table)
        {
            return CreateSubSectionPreviewTableCard(null, null, table);
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
            titleRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            titleRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var titleBlock = new TextBlock
            {
                Text = table.Title,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4")),
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(titleBlock, 0);
            titleRow.Children.Add(titleBlock);

            var toggleFiltersButton = CreateTableFiltersToggleButton(context);
            Grid.SetColumn(toggleFiltersButton, 1);
            titleRow.Children.Add(toggleFiltersButton);

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

                    var exportTable7Button = CreateSecondaryButton("Экспортировать по шаблону");
                    exportTable7Button.Tag = context;
                    exportTable7Button.Click += ExportToTable7Button_Click;
                    buttons.Children.Add(exportTable7Button);
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

        private Border CreateTableEditorCard(
            DynamicTemplateEntry templateEntry,
            SubSection subsection,
            Table table,
            ICollection<TableEditorContext> tableEditors = null)
        {
            var titleBox = CreateEditorTextBox(table.Title, false);
            var structure = CreateEditableTableStructure(table);
            var editorHost = new ContentControl();

            var context = new TableEditorContext
            {
                Template = templateEntry,
                SubSection = subsection,
                Table = table,
                TitleTextBox = titleBox,
                Structure = structure,
                TableEditorHost = editorHost,
                TableViewFactory = CreateEditableTableGrid
            };
            tableEditors?.Add(context);

            var panel = new StackPanel();
            panel.Children.Add(new TextBlock
            {
                Text = "Название таблицы",
                Margin = new Thickness(0, 0, 0, 6),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4"))
            });
            panel.Children.Add(titleBox);

            var tableActions = CreateButtonRow();

            var addColumnButton = CreateSecondaryButton("Добавить столбец");
            addColumnButton.Tag = context;
            addColumnButton.Click += AddTableColumnButton_Click;
            tableActions.Children.Add(addColumnButton);

            var addRowButton = CreateSecondaryButton("Добавить строку");
            addRowButton.Tag = context;
            addRowButton.Click += AddTableRowButton_Click;
            tableActions.Children.Add(addRowButton);

            var mergeCellsButton = CreateSecondaryButton("Объединить ячейки");
            mergeCellsButton.Tag = context;
            mergeCellsButton.Click += MergeTableCellsButton_Click;
            tableActions.Children.Add(mergeCellsButton);

            var importButton = CreateSecondaryButton("Импортировать таблицу");
            importButton.Tag = context;
            importButton.Click += ImportTableFromExcelButton_Click;
            tableActions.Children.Add(importButton);

            panel.Children.Add(tableActions);

            panel.Children.Add(new TextBlock
            {
                Text = "Нажмите на ячейку и печатайте прямо в таблице. Первая строка относится к шапке, новые строки добавляются как строки данных. Замочек рядом со строкой или столбцом переключает их в режим шапки и обратно.",
                Margin = new Thickness(0, 0, 0, 8),
                Foreground = Brushes.DimGray,
                TextWrapping = TextWrapping.Wrap
            });
            panel.Children.Add(new TextBlock
            {
                Text = "Редактор таблицы",
                Margin = new Thickness(0, 4, 0, 6),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4")),
                FontWeight = FontWeights.SemiBold
            });

            var editorScrollViewer = new ScrollViewer
            {
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxHeight = 420,
                Content = editorHost
            };
            AttachMouseWheelScrolling(editorScrollViewer);
            panel.Children.Add(editorScrollViewer);
            RefreshTableEditor(context);

            var buttons = CreateButtonRow();
            var deleteButton = CreateDangerButton("Удалить таблицу");
            deleteButton.Tag = context;
            deleteButton.Click += DeleteTableButton_Click;
            buttons.Children.Add(deleteButton);

            panel.Children.Add(buttons);

            return CreateContentCard($"Таблица: {table.Title}", panel);
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
            context.TableEditorHost.Content = (context.TableViewFactory ?? CreateEditableTableGrid)(context);
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

        private UIElement CreateEditableTableGrid(TableEditorContext context)
        {
            var structure = context.Structure;
            int totalDisplayRowCount = structure.HeaderRowCount + structure.BodyRowCount;

            var grid = new Grid
            {
                Background = Brushes.White
            };

            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(28) });
            for (int rowIndex = 0; rowIndex < totalDisplayRowCount; rowIndex++)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            }
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });

            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(28) });
            for (int columnIndex = 0; columnIndex < structure.ColumnCount; columnIndex++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition
                {
                    Width = new GridLength(160)
                });
            }
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(24) });

            AddColumnHandles(grid, context);
            AddRowHandles(grid, context);
            AddEditableCellsToGrid(grid, structure.HeaderCells, 1, 1);
            AddEditableCellsToGrid(grid, structure.BodyCells, structure.HeaderRowCount + 1, 1);

            return new Border
            {
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ced4da")),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                ClipToBounds = true,
                Child = grid
            };
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

            if (context?.FiltersEnabled == true)
            {
                panel.Children.Add(CreateTableColumnFiltersRow(context, columnCount));
            }

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
                    panel.Children.Add(CreateEmptyTableMessageBorder("По выбранным фильтрам строки не найдены."));
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

        private Button CreateTableFiltersToggleButton(TableEditorContext context)
        {
            bool isEnabled = context?.FiltersEnabled == true;
            var activeBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4"));
            var inactiveBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#d6e2ea"));
            var button = new Button
            {
                Width = 32,
                Height = 32,
                Margin = new Thickness(12, 0, 0, 0),
                Padding = new Thickness(0),
                BorderThickness = new Thickness(1),
                BorderBrush = isEnabled ? activeBrush : inactiveBrush,
                Background = isEnabled ? activeBrush : Brushes.White,
                Cursor = Cursors.Hand,
                ToolTip = isEnabled ? "Выключить фильтры таблицы" : "Включить фильтры таблицы",
                Content = CreateTableFilterIcon(isEnabled ? Brushes.White : activeBrush),
                Focusable = false
            };

            button.Click += (sender, args) =>
            {
                context.FiltersEnabled = !context.FiltersEnabled;
                if (!context.FiltersEnabled)
                {
                    context.ColumnFilters.Clear();
                }

                RefreshTableEditor(context);
            };

            return button;
        }

        private Button CreateDynamicSaveButton(Border ownerBorder)
        {
            var button = CreateDynamicIconButton("Сохранить в шаблоны");
            button.Content = CreateDynamicSaveIcon();
            button.PreviewMouseLeftButtonDown += (sender, args) => args.Handled = true;
            button.PreviewMouseLeftButtonUp += async (sender, args) =>
            {
                args.Handled = true;
                await SaveTemplateSnapshotAsync(ownerBorder);
            };

            return button;
        }

        private UIElement CreateTableColumnFiltersRow(TableEditorContext context, int columnCount)
        {
            var grid = new Grid
            {
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 0, 4)
            };
            ConfigurePreviewTableColumns(grid, columnCount);
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
            {
                bool isActive = TryGetTableColumnFilterState(context, columnIndex, out var filterState) && filterState.HasSettings;
                var border = new Border
                {
                    BorderBrush = Brushes.Black,
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Background = isActive
                        ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eef6ff"))
                        : Brushes.White,
                    Padding = new Thickness(6, 4, 6, 4),
                    Child = CreateTableColumnFilterButton(context, columnIndex, isActive)
                };

                Grid.SetColumn(border, columnIndex - 1);
                grid.Children.Add(border);
            }

            return grid;
        }

        private Button CreateTableColumnFilterButton(TableEditorContext context, int columnIndex, bool isActive)
        {
            var accentBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4"));
            var button = new Button
            {
                Width = 22,
                Height = 22,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Center,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                ToolTip = $"Фильтр столбца {columnIndex}",
                Content = CreateTableFilterIcon(isActive ? accentBrush : Brushes.DimGray),
                Focusable = false
            };

            button.Click += (sender, args) => ShowTableColumnFilterDialog(context, columnIndex, button);
            return button;
        }

        private UIElement CreateTableFilterIcon(Brush strokeBrush)
        {
            return new Viewbox
            {
                Width = 14,
                Height = 14,
                Child = new System.Windows.Shapes.Path
                {
                    Data = Geometry.Parse("M3 4.6C3 4.03995 3 3.75992 3.10899 3.54601C3.20487 3.35785 3.35785 3.20487 3.54601 3.10899C3.75992 3 4.03995 3 4.6 3H19.4C19.9601 3 20.2401 3 20.454 3.10899C20.6422 3.20487 20.7951 3.35785 20.891 3.54601C21 3.75992 21 4.03995 21 4.6V6.33726C21 6.58185 21 6.70414 20.9724 6.81923C20.9479 6.92127 20.9075 7.01881 20.8526 7.10828C20.7908 7.2092 20.7043 7.29568 20.5314 7.46863L14.4686 13.5314C14.2957 13.7043 14.2092 13.7908 14.1474 13.8917C14.0925 13.9812 14.0521 14.0787 14.0276 14.1808C14 14.2959 14 14.4182 14 14.6627V17L10 21V14.6627C10 14.4182 10 14.2959 9.97237 14.1808C9.94787 14.0787 9.90747 13.9812 9.85264 13.8917C9.7908 13.7908 9.70432 13.7043 9.53137 13.5314L3.46863 7.46863C3.29568 7.29568 3.2092 7.2092 3.14736 7.10828C3.09253 7.01881 3.05213 6.92127 3.02763 6.81923C3 6.70414 3 6.58185 3 6.33726V4.6Z"),
                    Stretch = Stretch.Uniform,
                    Stroke = strokeBrush,
                    StrokeThickness = 2,
                    StrokeStartLineCap = PenLineCap.Round,
                    StrokeEndLineCap = PenLineCap.Round,
                    StrokeLineJoin = PenLineJoin.Round
                }
            };
        }

        private void ShowTableColumnFilterDialog(TableEditorContext context, int columnIndex, FrameworkElement placementTarget)
        {
            TableColumnFilterState currentState = TryGetTableColumnFilterState(context, columnIndex, out var existingState)
                ? existingState
                : new TableColumnFilterState();
            TableColumnSortMode selectedSortMode = currentState.SortMode;

            var popup = new Popup
            {
                PlacementTarget = placementTarget,
                Placement = PlacementMode.Bottom,
                StaysOpen = false,
                AllowsTransparency = true,
                PopupAnimation = PopupAnimation.Fade
            };

            var accentBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4"));
            var borderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#d6e2ea"));

            var root = new Border
            {
                Width = 260,
                Background = Brushes.White,
                BorderBrush = borderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                SnapsToDevicePixels = true,
                Child = new StackPanel()
            };

            var panel = (StackPanel)root.Child;

            panel.Children.Add(new TextBlock
            {
                Text = "Поиск по значению",
                Margin = new Thickness(12, 10, 12, 6),
                Foreground = accentBrush,
                FontWeight = FontWeights.SemiBold
            });

            var searchBox = new TextBox
            {
                Text = currentState.SearchText ?? string.Empty,
                Margin = new Thickness(12, 0, 12, 10),
                Padding = new Thickness(8, 6, 8, 6)
            };
            panel.Children.Add(searchBox);

            Action closePopup = () =>
            {
                popup.IsOpen = false;
            };

            searchBox.KeyDown += (sender, args) =>
            {
                if (args.Key != Key.Enter)
                {
                    return;
                }

                ApplyTableColumnFilter(context, columnIndex, searchBox.Text, selectedSortMode);
                closePopup();
                args.Handled = true;
            };

            panel.Children.Add(new Border
            {
                Height = 1,
                Background = borderBrush
            });

            panel.Children.Add(CreateTableFilterPopupButton(
                "A↓",
                "Сортировка от А до Я",
                selectedSortMode == TableColumnSortMode.AlphabetAsc,
                () =>
                {
                    selectedSortMode = selectedSortMode == TableColumnSortMode.AlphabetAsc
                        ? TableColumnSortMode.None
                        : TableColumnSortMode.AlphabetAsc;
                    ApplyTableColumnFilter(context, columnIndex, searchBox.Text, selectedSortMode);
                    closePopup();
                }));

            panel.Children.Add(CreateTableFilterPopupButton(
                "A↑",
                "Сортировка от Я до А",
                selectedSortMode == TableColumnSortMode.AlphabetDesc,
                () =>
                {
                    selectedSortMode = selectedSortMode == TableColumnSortMode.AlphabetDesc
                        ? TableColumnSortMode.None
                        : TableColumnSortMode.AlphabetDesc;
                    ApplyTableColumnFilter(context, columnIndex, searchBox.Text, selectedSortMode);
                    closePopup();
                }));

            panel.Children.Add(CreateTableFilterPopupButton(
                "1↓",
                "Сортировка по возрастанию",
                selectedSortMode == TableColumnSortMode.ValueAsc,
                () =>
                {
                    selectedSortMode = selectedSortMode == TableColumnSortMode.ValueAsc
                        ? TableColumnSortMode.None
                        : TableColumnSortMode.ValueAsc;
                    ApplyTableColumnFilter(context, columnIndex, searchBox.Text, selectedSortMode);
                    closePopup();
                }));

            panel.Children.Add(CreateTableFilterPopupButton(
                "1↑",
                "Сортировка по убыванию",
                selectedSortMode == TableColumnSortMode.ValueDesc,
                () =>
                {
                    selectedSortMode = selectedSortMode == TableColumnSortMode.ValueDesc
                        ? TableColumnSortMode.None
                        : TableColumnSortMode.ValueDesc;
                    ApplyTableColumnFilter(context, columnIndex, searchBox.Text, selectedSortMode);
                    closePopup();
                }));

            if (currentState.HasSettings)
            {
                panel.Children.Add(new Border
                {
                    Height = 1,
                    Background = borderBrush
                });

                panel.Children.Add(CreateTableFilterPopupButton(
                    "×",
                    "Сбросить фильтр",
                    false,
                    () =>
                    {
                        context.ColumnFilters.Remove(columnIndex);
                        RefreshTableEditor(context);
                        closePopup();
                    }));
            }

            popup.Child = root;
            popup.IsOpen = true;
            searchBox.Focus();
            searchBox.SelectAll();
        }

        private Button CreateTableFilterPopupButton(string iconText, string text, bool isSelected, Action onClick)
        {
            var accentBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4"));
            var button = new Button
            {
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Padding = new Thickness(0),
                Background = isSelected
                    ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#eef6ff"))
                    : Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Focusable = false
            };

            var grid = new Grid
            {
                Margin = new Thickness(0)
            };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(34) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var iconBlock = new TextBlock
            {
                Text = iconText,
                Foreground = accentBrush,
                FontWeight = FontWeights.SemiBold,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(iconBlock, 0);
            grid.Children.Add(iconBlock);

            var textBlock = new TextBlock
            {
                Text = text,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#343a40")),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 8, 12, 8)
            };
            Grid.SetColumn(textBlock, 1);
            grid.Children.Add(textBlock);

            button.Content = grid;
            button.Click += (sender, args) => onClick();
            return button;
        }

        private void ApplyTableColumnFilter(TableEditorContext context, int columnIndex, string searchText, TableColumnSortMode sortMode)
        {
            string normalizedSearchText = (searchText ?? string.Empty).Trim();

            if (sortMode != TableColumnSortMode.None)
            {
                foreach (var otherKey in context.ColumnFilters.Keys.ToList())
                {
                    if (otherKey == columnIndex)
                    {
                        continue;
                    }

                    context.ColumnFilters[otherKey].SortMode = TableColumnSortMode.None;
                    if (!context.ColumnFilters[otherKey].HasSettings)
                    {
                        context.ColumnFilters.Remove(otherKey);
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(normalizedSearchText) && sortMode == TableColumnSortMode.None)
            {
                context.ColumnFilters.Remove(columnIndex);
            }
            else
            {
                context.ColumnFilters[columnIndex] = new TableColumnFilterState
                {
                    SearchText = normalizedSearchText,
                    SortMode = sortMode
                };
            }

            RefreshTableEditor(context);
        }

        private bool TryGetTableColumnFilterState(TableEditorContext context, int columnIndex, out TableColumnFilterState filterState)
        {
            filterState = null;
            return context != null
                && context.ColumnFilters != null
                && context.ColumnFilters.TryGetValue(columnIndex, out filterState)
                && filterState != null;
        }

        private IReadOnlyList<int> GetVisibleBodyRows(TableEditorContext context)
        {
            if (context?.Structure == null || context.Structure.BodyRowCount <= 0)
            {
                return Array.Empty<int>();
            }

            var rows = Enumerable.Range(1, context.Structure.BodyRowCount).ToList();
            if (context.ColumnFilters == null || context.ColumnFilters.Count == 0)
            {
                return rows;
            }

            foreach (var filter in context.ColumnFilters
                .Where(item => item.Value != null && !string.IsNullOrWhiteSpace(item.Value.SearchText))
                .OrderBy(item => item.Key))
            {
                string searchText = filter.Value.SearchText.Trim();
                rows = rows
                    .Where(row => GetBodyCellText(context.Structure, row, filter.Key)
                        .IndexOf(searchText, StringComparison.CurrentCultureIgnoreCase) >= 0)
                    .ToList();
            }

            var activeSort = context.ColumnFilters
                .Where(item => item.Value != null && item.Value.SortMode != TableColumnSortMode.None)
                .OrderBy(item => item.Key)
                .FirstOrDefault();

            if (activeSort.Value != null)
            {
                rows = rows
                    .OrderBy(row => row, Comparer<int>.Create((leftRow, rightRow) =>
                    {
                        int result = CompareBodyRows(
                            context.Structure,
                            leftRow,
                            rightRow,
                            activeSort.Key,
                            activeSort.Value.SortMode);
                        return result != 0 ? result : leftRow.CompareTo(rightRow);
                    }))
                    .ToList();
            }

            return rows;
        }

        private int CompareBodyRows(
            TableStructure structure,
            int leftRow,
            int rightRow,
            int columnIndex,
            TableColumnSortMode sortMode)
        {
            string leftText = GetBodyCellText(structure, leftRow, columnIndex);
            string rightText = GetBodyCellText(structure, rightRow, columnIndex);

            switch (sortMode)
            {
                case TableColumnSortMode.AlphabetAsc:
                    return StringComparer.CurrentCultureIgnoreCase.Compare(leftText, rightText);
                case TableColumnSortMode.AlphabetDesc:
                    return StringComparer.CurrentCultureIgnoreCase.Compare(rightText, leftText);
                case TableColumnSortMode.ValueAsc:
                    return CompareBodyRowValues(leftText, rightText);
                case TableColumnSortMode.ValueDesc:
                    return CompareBodyRowValues(rightText, leftText);
                default:
                    return 0;
            }
        }

        private int CompareBodyRowValues(string leftText, string rightText)
        {
            bool leftIsNumber = TryParseTableNumericValue(leftText, out decimal leftValue);
            bool rightIsNumber = TryParseTableNumericValue(rightText, out decimal rightValue);

            if (leftIsNumber && rightIsNumber)
            {
                return leftValue.CompareTo(rightValue);
            }

            if (leftIsNumber != rightIsNumber)
            {
                return leftIsNumber ? -1 : 1;
            }

            return StringComparer.CurrentCultureIgnoreCase.Compare(leftText, rightText);
        }

        private bool TryParseTableNumericValue(string value, out decimal numericValue)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (decimal.TryParse(normalized, NumberStyles.Any, CultureInfo.CurrentCulture, out numericValue))
            {
                return true;
            }

            return decimal.TryParse(normalized.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out numericValue);
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

        private string GetTemplateStorageDescription(DynamicTemplateEntry templateEntry)
        {
            return IsSharedTemplateDatabase(templateEntry.DatabasePath)
                ? $"Общая база шаблонов: {templateEntry.DatabasePath}"
                : $"База шаблона: {templateEntry.DatabasePath}";
        }

        private string GetSectionDisplayContent(Section section)
        {
            string content = GetSectionRawContent(section);
            return !string.IsNullOrWhiteSpace(content)
                ? content
                : "Текст раздела пока не заполнен.";
        }

        private string GetSectionRawContent(Section section)
        {
            var contentSubSection = GetSectionContentSubSection(section);
            return contentSubSection?.Texts != null && contentSubSection.Texts.Any()
                ? contentSubSection.Texts.First().Content
                : string.Empty;
        }

        private bool HasSectionDisplayContent(Section section)
        {
            var contentSubSection = GetSectionContentSubSection(section);
            bool hasText = !string.IsNullOrWhiteSpace(GetSectionRawContent(section));
            bool hasTables = contentSubSection?.Tables != null && contentSubSection.Tables.Any();
            return hasText || hasTables;
        }

        private bool HasSubSectionDisplayContent(SubSection subsection)
        {
            if (subsection == null)
            {
                return false;
            }

            bool hasText = subsection.Texts != null && subsection.Texts.Any(text =>
                !string.IsNullOrWhiteSpace(text?.Content));
            bool hasTables = subsection.Tables != null && subsection.Tables.Any();
            return hasText || hasTables;
        }

        private void AddColumnHandles(Grid grid, TableEditorContext context)
        {
            for (int column = 1; column <= context.Structure.ColumnCount; column++)
            {
                int targetColumn = column;
                bool canDelete = context.Structure.ColumnCount > 1;
                bool isHeaderColumn = IsTableColumnHeader(context.Structure, targetColumn);
                var handle = CreateTableEdgeHandle(
                    plusToolTip: "Вставить столбец",
                    minusToolTip: "Удалить столбец",
                    lockToolTip: isHeaderColumn ? "Убрать столбец из шапки" : "Сделать столбец шапкой",
                    onPlusClick: () =>
                    {
                        InsertTableColumn(context, targetColumn);
                    },
                    onMinusClick: canDelete
                        ? (System.Action)(() => DeleteTableColumn(context, targetColumn))
                        : null,
                    onLockClick: () =>
                    {
                        ToggleTableColumnHeader(context, targetColumn);
                    },
                    isLocked: isHeaderColumn,
                    isColumnHandle: true);

                Grid.SetRow(handle, 0);
                Grid.SetColumn(handle, targetColumn);
                grid.Children.Add(handle);
            }

            var trailingHandle = CreateTableEdgeHandle(
                plusToolTip: "Добавить столбец справа",
                minusToolTip: null,
                lockToolTip: null,
                onPlusClick: () =>
                {
                    InsertTableColumn(context, context.Structure.ColumnCount + 1);
                },
                onMinusClick: null,
                onLockClick: null,
                isLocked: false,
                isColumnHandle: true);

            Grid.SetRow(trailingHandle, 0);
            Grid.SetColumn(trailingHandle, context.Structure.ColumnCount + 1);
            grid.Children.Add(trailingHandle);
        }

        private void AddRowHandles(Grid grid, TableEditorContext context)
        {
            int totalStructureRowCount = context.Structure.HeaderRowCount + context.Structure.BodyRowCount;

            for (int structureRow = 1; structureRow <= totalStructureRowCount; structureRow++)
            {
                int targetRow = structureRow;
                bool isHeaderRow = structureRow <= context.Structure.HeaderRowCount;
                bool canDelete = isHeaderRow
                    ? context.Structure.HeaderRowCount > 1
                    : context.Structure.BodyRowCount > 0;
                bool rowMarkedAsHeader = IsDisplayRowHeader(context.Structure, targetRow);
                int gridRow = structureRow;

                var handle = CreateTableEdgeHandle(
                    plusToolTip: "Вставить строку",
                    minusToolTip: "Удалить строку",
                    lockToolTip: rowMarkedAsHeader ? "Убрать строку из шапки" : "Сделать строку шапкой",
                    onPlusClick: () =>
                    {
                        InsertTableRow(context, targetRow);
                    },
                    onMinusClick: canDelete
                        ? (System.Action)(() => DeleteTableRow(context, targetRow))
                        : null,
                    onLockClick: () =>
                    {
                        ToggleDisplayRowHeader(context, targetRow);
                    },
                    isLocked: rowMarkedAsHeader,
                    isColumnHandle: false);

                Grid.SetRow(handle, gridRow);
                Grid.SetColumn(handle, 0);
                grid.Children.Add(handle);
            }

            var trailingHandle = CreateTableEdgeHandle(
                plusToolTip: "Добавить строку снизу",
                minusToolTip: null,
                lockToolTip: null,
                onPlusClick: () =>
                {
                    InsertTableRow(context, totalStructureRowCount + 1);
                },
                onMinusClick: null,
                onLockClick: null,
                isLocked: false,
                isColumnHandle: false);

            Grid.SetRow(trailingHandle, totalStructureRowCount + 1);
            Grid.SetColumn(trailingHandle, 0);
            grid.Children.Add(trailingHandle);
        }

        private Border CreateTableEdgeHandle(
            string plusToolTip,
            string minusToolTip,
            string lockToolTip,
            System.Action onPlusClick,
            System.Action onMinusClick,
            System.Action onLockClick,
            bool isLocked,
            bool isColumnHandle)
        {
            var buttonsPanel = new StackPanel
            {
                Orientation = isColumnHandle ? Orientation.Horizontal : Orientation.Vertical,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Visibility = Visibility.Collapsed
            };

            if (onPlusClick != null)
            {
                buttonsPanel.Children.Add(CreateTableEdgeButton(CreateTablePlusIcon(), plusToolTip, onPlusClick));
            }

            if (onMinusClick != null)
            {
                buttonsPanel.Children.Add(CreateTableEdgeButton(CreateTableMinusIcon(), minusToolTip, onMinusClick));
            }

            if (onLockClick != null)
            {
                buttonsPanel.Children.Add(CreateTableEdgeButton(
                    isLocked ? CreateTableLockIcon() : CreateTableUnlockIcon(),
                    lockToolTip,
                    onLockClick));
            }

            var host = new Border
            {
                Background = Brushes.Transparent,
                Child = buttonsPanel
            };

            host.MouseEnter += (sender, args) =>
            {
                buttonsPanel.Visibility = Visibility.Visible;
            };
            host.MouseLeave += (sender, args) =>
            {
                buttonsPanel.Visibility = Visibility.Collapsed;
            };

            return host;
        }

        private Button CreateTableEdgeButton(UIElement icon, string toolTip, System.Action onClick)
        {
            var button = new Button
            {
                Width = 16,
                Height = 16,
                Margin = new Thickness(2),
                Padding = new Thickness(0),
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                ToolTip = toolTip,
                Content = icon,
                Focusable = false
            };

            button.Click += (sender, args) => onClick();
            return button;
        }

        private UIElement CreateTablePlusIcon()
        {
            var grid = new Grid
            {
                Width = 14,
                Height = 14
            };

            grid.Children.Add(new System.Windows.Shapes.Ellipse
            {
                Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4CAF50"))
            });

            grid.Children.Add(new Border
            {
                Width = 8,
                Height = 2,
                Background = Brushes.White,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            });

            grid.Children.Add(new Border
            {
                Width = 2,
                Height = 8,
                Background = Brushes.White,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            });

            return grid;
        }

        private UIElement CreateTableMinusIcon()
        {
            var grid = new Grid
            {
                Width = 14,
                Height = 14
            };

            grid.Children.Add(new Border
            {
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF3F62")),
                CornerRadius = new CornerRadius(3)
            });

            grid.Children.Add(new Border
            {
                Width = 8,
                Height = 2,
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#830018")),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            });

            return grid;
        }

        private UIElement CreateTableUnlockIcon()
        {
            var grid = new Grid
            {
                Width = 14,
                Height = 14
            };

            grid.Children.Add(new Border
            {
                Width = 10,
                Height = 7,
                CornerRadius = new CornerRadius(2),
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#111827")),
                VerticalAlignment = VerticalAlignment.Bottom
            });

            grid.Children.Add(new System.Windows.Shapes.Path
            {
                Stroke = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#111827")),
                StrokeThickness = 1.8,
                StrokeStartLineCap = PenLineCap.Round,
                StrokeEndLineCap = PenLineCap.Round,
                Data = Geometry.Parse("M4.3,7 L4.3,5.2 C4.3,3.3 5.6,2 7,2 C8.4,2 9.6,3.2 9.6,4.8"),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Top
            });

            grid.Children.Add(new Border
            {
                Width = 2,
                Height = 2,
                Background = Brushes.White,
                CornerRadius = new CornerRadius(1),
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, 0, 3)
            });

            return grid;
        }

        private UIElement CreateTableLockIcon()
        {
            return new Viewbox
            {
                Width = 14,
                Height = 14,
                Stretch = Stretch.Uniform,
                Child = new System.Windows.Shapes.Path
                {
                    Fill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#111827")),
                    Data = Geometry.Parse("M 12 1 C 8.6761905 1 6 3.6761905 6 7 L 6 8 C 4.9 8 4 8.9 4 10 L 4 20 C 4 21.1 4.9 22 6 22 L 18 22 C 19.1 22 20 21.1 20 20 L 20 10 C 20 8.9 19.1 8 18 8 L 18 7 C 18 3.6761905 15.32381 1 12 1 z M 12 3 C 14.27619 3 16 4.7238095 16 7 L 16 8 L 8 8 L 8 7 C 8 4.7238095 9.7238095 3 12 3 z M 12 13 C 13.1 13 14 13.9 14 15 C 14 16.1 13.1 17 12 17 C 10.9 17 10 16.1 10 15 C 10 13.9 10.9 13 12 13 z")
                }
            };
        }

        private void AddEditableCellsToGrid(Grid grid, IEnumerable<TableCellDefinition> cells, int rowOffset, int columnOffset)
        {
            foreach (var cell in cells.OrderBy(item => item.Row).ThenBy(item => item.Column))
            {
                var textBox = CreateEditableTableCellTextBox(cell);
                Grid.SetRow(textBox, rowOffset + cell.Row - 1);
                Grid.SetColumn(textBox, columnOffset + cell.Column - 1);
                Grid.SetColumnSpan(textBox, cell.ColSpan);
                Grid.SetRowSpan(textBox, cell.RowSpan);
                grid.Children.Add(textBox);
            }
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
            foreach (var seed in StudyReportSubSections)
            {
                if (string.IsNullOrWhiteSpace(seed?.TablePatternName) || seed.Headers == null)
                {
                    continue;
                }

                yield return new TableSeed
                {
                    Title = seed.Title,
                    PatternName = seed.TablePatternName,
                    Cells = seed.Headers
                        .Select((header, index) => HeaderCell(1, index + 1, header))
                        .ToArray()
                };
            }

            foreach (var section in GetNirReportSections())
            {
                foreach (var tableSeed in EnumerateTableSeeds(section))
                {
                    yield return tableSeed;
                }
            }
        }

        private static IEnumerable<TableSeed> EnumerateTableSeeds(NirSectionSeed section)
        {
            foreach (var table in section?.Tables ?? Array.Empty<TableSeed>())
            {
                yield return table;
            }

            foreach (var subsection in section?.SubSections ?? Array.Empty<NirSubSectionSeed>())
            {
                foreach (var table in EnumerateTableSeeds(subsection))
                {
                    yield return table;
                }
            }
        }

        private static IEnumerable<TableSeed> EnumerateTableSeeds(NirSubSectionSeed subsection)
        {
            foreach (var table in subsection?.Tables ?? Array.Empty<TableSeed>())
            {
                yield return table;
            }

            foreach (var child in subsection?.Children ?? Array.Empty<NirSubSectionSeed>())
            {
                foreach (var table in EnumerateTableSeeds(child))
                {
                    yield return table;
                }
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

        private void ToggleDisplayRowHeader(TableEditorContext context, int displayRow)
        {
            EnsureEditableTableStructure(context.Structure);

            var rowCells = GetCellsForDisplayRow(context.Structure, displayRow).ToList();
            if (!rowCells.Any())
            {
                return;
            }

            bool targetState = !rowCells.All(cell => cell.IsHeader);
            foreach (var cell in rowCells)
            {
                cell.IsHeader = targetState;
            }

            RefreshTableEditor(context);
        }

        private void ToggleTableColumnHeader(TableEditorContext context, int column)
        {
            EnsureEditableTableStructure(context.Structure);

            var columnCells = GetCellsForColumn(context.Structure, column).ToList();
            if (!columnCells.Any())
            {
                return;
            }

            bool targetState = !columnCells.All(cell => cell.IsHeader);
            foreach (var cell in columnCells)
            {
                cell.IsHeader = targetState;
            }

            RefreshTableEditor(context);
        }

        private bool IsDisplayRowHeader(TableStructure structure, int displayRow)
        {
            var rowCells = GetCellsForDisplayRow(structure, displayRow).ToList();
            return rowCells.Any() && rowCells.All(cell => cell.IsHeader);
        }

        private bool IsTableColumnHeader(TableStructure structure, int column)
        {
            var columnCells = GetCellsForColumn(structure, column).ToList();
            return columnCells.Any() && columnCells.All(cell => cell.IsHeader);
        }

        private IEnumerable<TableCellDefinition> GetCellsForDisplayRow(TableStructure structure, int displayRow)
        {
            if (structure == null)
            {
                yield break;
            }

            if (displayRow <= structure.HeaderRowCount)
            {
                foreach (var cell in structure.HeaderCells.Where(cell =>
                             cell.Row <= displayRow && displayRow < cell.Row + cell.RowSpan))
                {
                    yield return cell;
                }

                yield break;
            }

            int bodyRow = displayRow - structure.HeaderRowCount;
            foreach (var cell in structure.BodyCells.Where(cell =>
                         cell.Row <= bodyRow && bodyRow < cell.Row + cell.RowSpan))
            {
                yield return cell;
            }
        }

        private IEnumerable<TableCellDefinition> GetCellsForColumn(TableStructure structure, int column)
        {
            if (structure == null)
            {
                yield break;
            }

            foreach (var cell in structure.HeaderCells.Concat(structure.BodyCells).Where(cell =>
                         cell.Column <= column && column < cell.Column + cell.ColSpan))
            {
                yield return cell;
            }
        }

        private void AddTableColumnButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

            InsertTableColumn(context, context.Structure.ColumnCount + 1);
        }

        private void AddTableRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

            InsertTableRow(context, context.Structure.HeaderRowCount + context.Structure.BodyRowCount + 1);
        }

        private void InsertTableColumn(TableEditorContext context, int insertAtColumn)
        {
            EnsureEditableTableStructure(context.Structure);
            ShiftCellsForInsertedColumn(context.Structure.HeaderCells, insertAtColumn);
            ShiftCellsForInsertedColumn(context.Structure.BodyCells, insertAtColumn);
            context.Structure.ColumnCount++;
            RefreshTableEditor(context);
        }

        private void DeleteTableColumn(TableEditorContext context, int deleteColumn)
        {
            EnsureEditableTableStructure(context.Structure);
            if (context.Structure.ColumnCount <= 1)
            {
                return;
            }

            ShiftCellsForDeletedColumn(context.Structure.HeaderCells, deleteColumn);
            ShiftCellsForDeletedColumn(context.Structure.BodyCells, deleteColumn);
            context.Structure.ColumnCount--;
            RefreshTableEditor(context);
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

        private void DeleteTableRow(TableEditorContext context, int visualRow)
        {
            EnsureEditableTableStructure(context.Structure);

            if (visualRow <= context.Structure.HeaderRowCount)
            {
                if (context.Structure.HeaderRowCount <= 1)
                {
                    return;
                }

                ShiftCellsForDeletedRow(context.Structure.HeaderCells, visualRow);
                context.Structure.HeaderRowCount--;
            }
            else
            {
                if (context.Structure.BodyRowCount <= 0)
                {
                    return;
                }

                int bodyRow = visualRow - context.Structure.HeaderRowCount;
                ShiftCellsForDeletedRow(context.Structure.BodyCells, bodyRow);
                context.Structure.BodyRowCount--;
            }

            RefreshTableEditor(context);
        }

        private void ShiftCellsForInsertedColumn(List<TableCellDefinition> cells, int insertAtColumn)
        {
            foreach (var cell in cells)
            {
                int cellEndColumn = cell.Column + cell.ColSpan - 1;
                if (cell.Column >= insertAtColumn)
                {
                    cell.Column++;
                }
                else if (cell.Column < insertAtColumn && cellEndColumn >= insertAtColumn)
                {
                    cell.ColSpan++;
                }
            }
        }

        private void ShiftCellsForDeletedColumn(List<TableCellDefinition> cells, int deleteColumn)
        {
            foreach (var cell in cells.ToList())
            {
                int cellEndColumn = cell.Column + cell.ColSpan - 1;
                if (cell.Column > deleteColumn)
                {
                    cell.Column--;
                    continue;
                }

                if (cell.Column <= deleteColumn && cellEndColumn >= deleteColumn)
                {
                    if (cell.ColSpan > 1)
                    {
                        cell.ColSpan--;
                    }
                    else
                    {
                        cells.Remove(cell);
                    }
                }
            }
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

        private void ShiftCellsForDeletedRow(List<TableCellDefinition> cells, int deleteRow)
        {
            foreach (var cell in cells.ToList())
            {
                int cellEndRow = cell.Row + cell.RowSpan - 1;
                if (cell.Row > deleteRow)
                {
                    cell.Row--;
                    continue;
                }

                if (cell.Row <= deleteRow && cellEndRow >= deleteRow)
                {
                    if (cell.RowSpan > 1)
                    {
                        cell.RowSpan--;
                    }
                    else
                    {
                        cells.Remove(cell);
                    }
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

        private async void AddSectionButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is DynamicTemplateEntry templateEntry))
            {
                return;
            }

            var database = new DataBase(templateEntry.DatabasePath);
            int nextNumber = templateEntry.Report.Sections.Any()
                ? templateEntry.Report.Sections.Max(item => item.Number) + 1
                : 1;

            int sectionId = await database.AddSection(new Section
            {
                ReportId = templateEntry.ReportId,
                Number = nextNumber,
                Title = GetDefaultSectionTitle(nextNumber)
            });

            await LogHistoryAsync(
                HistoryActionCreate,
                "section",
                $"{BuildTemplateHistoryLocation(templateEntry)} / {GetDefaultSectionTitle(nextNumber)}",
                "Создан раздел");
            await RefreshTemplateEntryAsync(templateEntry, sectionId: sectionId);
        }

        private async void AddSubSectionButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button))
            {
                return;
            }

            DynamicTemplateEntry templateEntry;
            Section section;
            SubSection parentSubSection;
            IEnumerable<SubSection> siblingSubSections;

            if (button.Tag is SectionEditorContext sectionContext)
            {
                templateEntry = sectionContext.Template;
                section = sectionContext.Section;
                parentSubSection = null;
                siblingSubSections = GetVisibleSubSections(sectionContext.Section);
            }
            else if (button.Tag is SubSectionEditorContext subSectionContext)
            {
                templateEntry = subSectionContext.Template;
                section = subSectionContext.SubSection.Section;
                parentSubSection = subSectionContext.SubSection;

                if (parentSubSection.ParentSubsectionId.HasValue)
                {
                    MessageBox.Show(
                        "Во вложенном подразделе нельзя создавать дополнительные вложенные подразделы.",
                        "Подраздел",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                siblingSubSections = subSectionContext.SubSection.SubSections ?? Enumerable.Empty<SubSection>();
            }
            else
            {
                return;
            }

            var dialog = new TemplateNameDialog
            {
                Owner = this,
                Title = "Новый подраздел",
                Prompt = "Введите название подраздела",
                Label = "Название подраздела"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            var database = new DataBase(templateEntry.DatabasePath);
            int nextNumber = siblingSubSections.Any()
                ? siblingSubSections.Max(item => item.Number) + 1
                : 1;

            int subsectionId = await database.AddSubsection(new SubSection
            {
                SectionId = section.Id,
                ParentSubsectionId = parentSubSection?.Id,
                Number = nextNumber,
                Title = dialog.TemplateName.Trim()
            });

            var createdSubSection = new SubSection
            {
                Id = subsectionId,
                Section = section,
                SectionId = section.Id,
                ParentSubsection = parentSubSection,
                ParentSubsectionId = parentSubSection?.Id,
                Number = nextNumber,
                Title = dialog.TemplateName.Trim()
            };
            await LogHistoryAsync(
                HistoryActionCreate,
                "subsection",
                BuildSubSectionHistoryLocation(templateEntry, createdSubSection),
                parentSubSection == null ? "Создан подраздел" : "Создан вложенный подраздел");
            await RefreshTemplateEntryAsync(templateEntry, subsectionId: subsectionId);
        }

        private async Task<SubSection> EnsureSectionContentSubSectionAsync(DynamicTemplateEntry templateEntry, Section section)
        {
            var existingSubSection = GetSectionContentSubSection(section);
            if (existingSubSection != null)
            {
                return existingSubSection;
            }

            var database = new DataBase(templateEntry.DatabasePath);
            int subsectionId = await database.AddSubsection(new SubSection
            {
                SectionId = section.Id,
                ParentSubsectionId = null,
                Number = 0,
                Title = SectionContentSubSectionTitle
            });

            var newSubSection = new SubSection
            {
                Id = subsectionId,
                SectionId = section.Id,
                ParentSubsectionId = null,
                Number = 0,
                Title = SectionContentSubSectionTitle,
                Section = section,
                SubSections = new List<SubSection>(),
                Tables = new List<Table>(),
                Texts = new List<Text>()
            };

            if (section.SubSections == null)
            {
                section.SubSections = new List<SubSection>();
            }

            section.SubSections.Add(newSubSection);
            return newSubSection;
        }

        private async void SaveSectionButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is SectionEditorContext context))
            {
                return;
            }

            string title = context.TitleTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(title))
            {
                MessageBox.Show("Название раздела не может быть пустым.", "Раздел",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string previousTitle = context.Section.Title?.Trim() ?? string.Empty;
            string previousContent = GetSectionRawContent(context.Section)?.Trim() ?? string.Empty;

            foreach (var tableContext in context.TableEditors)
            {
                if (!await SaveTableAsync(tableContext, false))
                {
                    return;
                }
            }

            context.Section.Title = title;
            var database = new DataBase(context.Template.DatabasePath);
            await database.UpdateSection(context.Section);
            string content = context.ContentTextBox?.Text.Trim() ?? string.Empty;
            var contentSubSection = context.ContentSubSection;

            if (!string.IsNullOrWhiteSpace(content))
            {
                contentSubSection = await EnsureSectionContentSubSectionAsync(context.Template, context.Section);
                Text existingText = contentSubSection.Texts.FirstOrDefault();
                if (existingText == null)
                {
                    await database.AddText(new Text
                    {
                        SubsectionId = contentSubSection.Id,
                        Content = content,
                        PatternName = "main_text"
                    });
                }
                else
                {
                    existingText.Content = content;
                    await database.UpdateText(existingText);
                }
            }
            else if (contentSubSection != null)
            {
                Text existingText = contentSubSection.Texts.FirstOrDefault();
                if (existingText != null)
                {
                    await database.DeleteText(existingText.Id);
                }
            }

            bool titleChanged = !string.Equals(previousTitle, title, StringComparison.Ordinal);
            bool contentChanged = !string.Equals(previousContent, content, StringComparison.Ordinal);
            if (titleChanged || contentChanged)
            {
                await LogHistoryAsync(
                    HistoryActionEdit,
                    "section",
                    BuildSectionHistoryLocation(context.Template, context.Section),
                    BuildHistoryDetails(
                        titleChanged ? "обновлено название" : null,
                        contentChanged ? "обновлен текст" : null));
            }

            await RefreshTemplateEntryAsync(context.Template, sectionId: context.Section.Id, editMode: false);
        }

        private async void SaveSubSectionButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is SubSectionEditorContext context))
            {
                return;
            }

            string title = context.TitleTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(title))
            {
                MessageBox.Show("Название подраздела не может быть пустым.", "Подраздел",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string previousTitle = context.SubSection.Title?.Trim() ?? string.Empty;
            string previousContent = context.SubSection.Texts != null && context.SubSection.Texts.Any()
                ? context.SubSection.Texts.First().Content?.Trim() ?? string.Empty
                : string.Empty;

            foreach (var tableContext in context.TableEditors)
            {
                if (!await SaveTableAsync(tableContext, false))
                {
                    return;
                }
            }

            string content = context.ContentTextBox.Text.Trim();
            var database = new DataBase(context.Template.DatabasePath);

            context.SubSection.Title = title;
            await database.UpdateSubsection(context.SubSection);

            Text existingText = context.SubSection.Texts.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(content))
            {
                if (existingText != null)
                {
                    await database.DeleteText(existingText.Id);
                }
            }
            else if (existingText == null)
            {
                await database.AddText(new Text
                {
                    SubsectionId = context.SubSection.Id,
                    Content = content,
                    PatternName = "main_text"
                });
            }
            else
            {
                existingText.Content = content;
                await database.UpdateText(existingText);
            }

            bool titleChanged = !string.Equals(previousTitle, title, StringComparison.Ordinal);
            bool contentChanged = !string.Equals(previousContent, content, StringComparison.Ordinal);
            if (titleChanged || contentChanged)
            {
                await LogHistoryAsync(
                    HistoryActionEdit,
                    "subsection",
                    BuildSubSectionHistoryLocation(context.Template, context.SubSection),
                    BuildHistoryDetails(
                        titleChanged ? "обновлено название" : null,
                        contentChanged ? "обновлен текст" : null));
            }

            await RefreshTemplateEntryAsync(context.Template, subsectionId: context.SubSection.Id, editMode: false);
        }

        private async void RenameTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is DynamicTemplateEntry templateEntry))
            {
                return;
            }

            var dialog = new TemplateNameDialog
            {
                Owner = this,
                Title = "Переименовать шаблон",
                Prompt = "Введите новое название шаблона",
                Label = "Название шаблона",
                TemplateName = templateEntry.DisplayTitle
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            string newTitle = dialog.TemplateName.Trim();
            if (string.IsNullOrWhiteSpace(newTitle))
            {
                return;
            }

            string previousTitle = templateEntry.DisplayTitle?.Trim() ?? string.Empty;
            if (string.Equals(previousTitle, newTitle, StringComparison.Ordinal))
            {
                return;
            }

            var database = new DataBase(templateEntry.DatabasePath);
            templateEntry.Report.Title = newTitle;
            templateEntry.Report.PatternFile.Title = newTitle;

            await database.UpdateReport(templateEntry.Report);
            await database.UpdateFilePattern(templateEntry.Report.PatternFile);

            templateEntry.DisplayTitle = newTitle;
            await LogHistoryAsync(
                HistoryActionEdit,
                "template",
                BuildTemplateHistoryLocation(templateEntry),
                $"переименован из \"{previousTitle}\"");
            await RefreshTemplateEntryAsync(templateEntry);
        }

        private async void DeleteTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is DynamicTemplateEntry templateEntry))
            {
                return;
            }

            await DeleteTemplateAsync(templateEntry);
        }

        private async void DeleteSectionButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is SectionEditorContext context))
            {
                return;
            }

            await DeleteSectionAsync(context.Template, context.Section);
        }

        private async void DeleteSubSectionButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is SubSectionEditorContext context))
            {
                return;
            }

            await DeleteSubSectionAsync(context.Template, context.SubSection);
        }

        private async Task DeleteTemplateAsync(DynamicTemplateEntry templateEntry)
        {
            if (templateEntry == null)
            {
                return;
            }

            if (MessageBox.Show($"Удалить шаблон \"{templateEntry.DisplayTitle}\"?",
                "Шаблоны", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            try
            {
                var database = new DataBase(templateEntry.DatabasePath);

                if (IsSharedTemplateDatabase(templateEntry.DatabasePath))
                {
                    int patternId = templateEntry.Report?.PatternFile?.Id ?? templateEntry.Report.PattarnId;
                    await database.DeleteFilePattern(patternId);
                }
                else if (File.Exists(templateEntry.DatabasePath))
                {
                    File.Delete(templateEntry.DatabasePath);
                }

                await LogHistoryAsync(
                    HistoryActionDelete,
                    "template",
                    BuildTemplateHistoryLocation(templateEntry),
                    "Удален шаблон");
                RemoveTemplateFromMenu(templateEntry);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось удалить шаблон:{Environment.NewLine}{ex.Message}",
                    "Шаблоны", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            MainContentControl.Content = new TextBlock
            {
                Text = "Шаблон удален",
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = Brushes.Gray,
                FontSize = 16
            };
            currentDynamicEditorBorder = null;
        }

        private async Task DeleteSectionAsync(DynamicTemplateEntry templateEntry, Section section)
        {
            if (templateEntry == null || section == null)
            {
                return;
            }

            if (MessageBox.Show($"Удалить раздел \"{BuildSectionTitle(section)}\"?",
                "Раздел", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            var database = new DataBase(templateEntry.DatabasePath);
            await database.DeleteSection(section.Id);
            await LogHistoryAsync(
                HistoryActionDelete,
                "section",
                BuildSectionHistoryLocation(templateEntry, section),
                "Удален раздел");
            await RefreshTemplateEntryAsync(templateEntry);
        }

        private async Task DeleteSubSectionAsync(DynamicTemplateEntry templateEntry, SubSection subsection)
        {
            if (templateEntry == null || subsection == null)
            {
                return;
            }

            if (MessageBox.Show($"Удалить подраздел \"{BuildSubSectionTitle(subsection.Section, subsection)}\"?",
                "Подраздел", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            var database = new DataBase(templateEntry.DatabasePath);
            await database.DeleteSubsection(subsection.Id);
            await LogHistoryAsync(
                HistoryActionDelete,
                "subsection",
                BuildSubSectionHistoryLocation(templateEntry, subsection),
                subsection.ParentSubsectionId.HasValue ? "Удален вложенный подраздел" : "Удален подраздел");

            if (subsection.ParentSubsectionId.HasValue)
            {
                await RefreshTemplateEntryAsync(templateEntry, subsectionId: subsection.ParentSubsectionId.Value);
            }
            else
            {
                await RefreshTemplateEntryAsync(templateEntry, sectionId: subsection.SectionId);
            }
        }

        private async void AddTableButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button))
            {
                return;
            }

            var dialog = new TemplateNameDialog
            {
                Owner = this,
                Title = "Новая таблица",
                Prompt = "Введите название таблицы",
                Label = "Название таблицы"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            DynamicTemplateEntry templateEntry;
            SubSection targetSubSection;
            int refreshSubSectionId;
            int? refreshSectionId = null;

            if (button.Tag is SubSectionEditorContext subSectionContext)
            {
                templateEntry = subSectionContext.Template;
                targetSubSection = subSectionContext.SubSection;
                refreshSubSectionId = subSectionContext.SubSection.Id;
            }
            else if (button.Tag is SectionEditorContext sectionContext)
            {
                templateEntry = sectionContext.Template;
                targetSubSection = await EnsureSectionContentSubSectionAsync(sectionContext.Template, sectionContext.Section);
                refreshSubSectionId = targetSubSection.Id;
                refreshSectionId = sectionContext.Section.Id;
            }
            else
            {
                return;
            }

            var database = new DataBase(templateEntry.DatabasePath);
            int tableId = await database.AddTable(new Table
            {
                Title = dialog.TemplateName.Trim(),
                SubsectionId = targetSubSection.Id,
                PatternName = $"table_{DateTime.Now.Ticks}"
            });

            await database.AddTableItem(new TableItem
            {
                TableId = tableId,
                Row = 1,
                Column = 1,
                Header = string.Empty,
                ColSpan = 1,
                RowSpan = 1,
                IsHeader = true
            });

            await LogHistoryAsync(
                HistoryActionCreate,
                "table",
                BuildTableHistoryLocation(templateEntry, targetSubSection, new Table { Title = dialog.TemplateName.Trim() }),
                "Создана таблица");
            if (refreshSectionId.HasValue)
            {
                await RefreshTemplateEntryAsync(templateEntry, sectionId: refreshSectionId.Value);
            }
            else
            {
                await RefreshTemplateEntryAsync(templateEntry, subsectionId: refreshSubSectionId);
            }
        }

        private async void SaveTableButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

            if (await SaveTableAsync(context, true))
            {
                MessageBox.Show("Таблица сохранена.", "Таблица",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
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

        private void ExportTableToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

            try
            {
                EnsureEditableTableStructure(context.Structure);

                string title = (context.TitleTextBox?.Text ?? context.Table?.Title ?? "Таблица").Trim();
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = "Таблица";
                }

                var saveDialog = new SaveFileDialog
                {
                    Title = "Экспортировать таблицу",
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    FileName = $"{SanitizeFileName(title)}.xlsx",
                    DefaultExt = ".xlsx",
                    AddExtension = true
                };

                if (saveDialog.ShowDialog() != true)
                {
                    return;
                }

                ExcelTableExchangeService.Export(saveDialog.FileName, ConvertToExcelTableData(context.Structure));
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

        private void ExportToTable7Button_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

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
                    ExportToStudyPublishingPlan(context);
                    return;
                }

                ExportToPublicationsListTemplate(context);
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

        private void ExportToStudyPublishingPlan(TableEditorContext context)
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

        private void ExportToPublicationsListTemplate(TableEditorContext context)
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
            context.ColumnFilters.Clear();

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
            context.ColumnFilters.Clear();

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

        private ExcelTableData ConvertToExcelTableData(TableStructure structure)
        {
            var data = new ExcelTableData
            {
                ColumnCount = structure?.ColumnCount ?? 0,
                HeaderRowCount = structure?.HeaderRowCount ?? 0,
                BodyRowCount = structure?.BodyRowCount ?? 0
            };

            if (structure == null)
            {
                return data;
            }

            foreach (var cell in structure.HeaderCells)
            {
                data.HeaderCells.Add(new ExcelTableCell
                {
                    Text = cell.Text,
                    Column = cell.Column,
                    Row = cell.Row,
                    ColSpan = cell.ColSpan,
                    RowSpan = cell.RowSpan,
                    IsHeader = cell.IsHeader
                });
            }

            foreach (var cell in structure.BodyCells)
            {
                data.BodyCells.Add(new ExcelTableCell
                {
                    Text = cell.Text,
                    Column = cell.Column,
                    Row = cell.Row,
                    ColSpan = cell.ColSpan,
                    RowSpan = cell.RowSpan,
                    IsHeader = cell.IsHeader
                });
            }

            return data;
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

        private void MergeTableCellsButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

            EnsureEditableTableStructure(context.Structure);
            var structure = context.Structure;
            if (!structure.HeaderCells.Any() && !structure.BodyCells.Any())
            {
                MessageBox.Show("Сначала заполните таблицу.", "Таблица",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (ShowMergeTableCellsDialog(structure))
            {
                RefreshTableEditor(context);
            }
        }

        private async void DeleteTableButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || !(button.Tag is TableEditorContext context))
            {
                return;
            }

            if (MessageBox.Show($"Удалить таблицу \"{context.Table.Title}\"?",
                "Таблица", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }

            var database = new DataBase(context.Template.DatabasePath);
            await database.DeleteTable(context.Table.Id);
            await LogHistoryAsync(
                HistoryActionDelete,
                "table",
                BuildTableHistoryLocation(context.Template, context.SubSection, context.Table),
                "Удалена таблица");
            await RefreshTemplateEntryAsync(context.Template, subsectionId: context.SubSection.Id);
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

        private bool ShowMergeTableCellsDialog(TableStructure structure)
        {
            var selectableCells = structure.HeaderCells
                .Concat(structure.BodyCells)
                .Where(cell => cell.ColSpan == 1 && cell.RowSpan == 1)
                .ToList();

            if (selectableCells.Count < 2)
            {
                MessageBox.Show("Для объединения нужны как минимум две обычные соседние ячейки таблицы.",
                    "Таблица", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            var window = new Window
            {
                Owner = this,
                Title = "Объединить ячейки",
                Width = 720,
                Height = 520,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                Background = Brushes.White
            };

            var root = new Grid
            {
                Margin = new Thickness(18)
            };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            root.Children.Add(new TextBlock
            {
                Text = "Выберите минимум две соседние ячейки таблицы и нажмите ОК.",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 12),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#495057"))
            });

            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
            };
            Grid.SetRow(scrollViewer, 1);

            var grid = new Grid
            {
                Background = Brushes.White
            };

            for (int columnIndex = 0; columnIndex < structure.ColumnCount; columnIndex++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition
                {
                    Width = columnIndex == 0 ? new GridLength(70) : new GridLength(130)
                });
            }

            int totalRowCount = structure.HeaderRowCount + structure.BodyRowCount;
            for (int rowIndex = 0; rowIndex < totalRowCount; rowIndex++)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(72) });
            }

            var buttonMap = new Dictionary<ToggleButton, TableCellDefinition>();
            var occupied = new bool[totalRowCount + 1, structure.ColumnCount + 1];

            foreach (var cell in structure.HeaderCells
                .Concat(structure.BodyCells)
                .OrderBy(c => GetDisplayRow(structure, c))
                .ThenBy(c => c.Column))
            {
                int displayRow = GetDisplayRow(structure, cell);
                var toggle = new ToggleButton
                {
                    Content = new TextBlock
                    {
                        Text = string.IsNullOrWhiteSpace(cell.Text) ? "(пусто)" : cell.Text,
                        TextWrapping = TextWrapping.Wrap,
                        TextAlignment = TextAlignment.Center,
                        Foreground = Brushes.Black,
                        FontWeight = cell.IsHeader ? FontWeights.Bold : FontWeights.Normal
                    },
                    Margin = new Thickness(0),
                    Padding = new Thickness(8),
                    BorderBrush = Brushes.Black,
                    BorderThickness = new Thickness(1),
                    Background = cell.IsHeader
                        ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f8f9fa"))
                        : Brushes.White,
                    Focusable = false,
                    IsEnabled = cell.ColSpan == 1 && cell.RowSpan == 1
                };

                if (!toggle.IsEnabled)
                {
                    toggle.ToolTip = "Эта ячейка уже объединена.";
                    toggle.Opacity = 0.7;
                }

                toggle.Checked += (checkedSender, args) =>
                {
                    ((ToggleButton)checkedSender).Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#dce8ff"));
                };
                toggle.Unchecked += (uncheckedSender, args) =>
                {
                    ((ToggleButton)uncheckedSender).Background = cell.IsHeader
                        ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f8f9fa"))
                        : Brushes.White;
                };

                Grid.SetRow(toggle, displayRow - 1);
                Grid.SetColumn(toggle, cell.Column - 1);
                Grid.SetColumnSpan(toggle, cell.ColSpan);
                Grid.SetRowSpan(toggle, cell.RowSpan);
                grid.Children.Add(toggle);

                if (toggle.IsEnabled)
                {
                    buttonMap[toggle] = cell;
                }

                for (int row = displayRow; row < displayRow + cell.RowSpan; row++)
                {
                    for (int column = cell.Column; column < cell.Column + cell.ColSpan; column++)
                    {
                        occupied[row, column] = true;
                    }
                }
            }

            for (int row = 1; row <= totalRowCount; row++)
            {
                for (int column = 1; column <= structure.ColumnCount; column++)
                {
                    if (occupied[row, column])
                    {
                        continue;
                    }

                    var filler = new Border
                    {
                        BorderBrush = Brushes.Black,
                        BorderThickness = new Thickness(1),
                        Background = Brushes.White
                    };
                    Grid.SetRow(filler, row - 1);
                    Grid.SetColumn(filler, column - 1);
                    grid.Children.Add(filler);
                }
            }

            scrollViewer.Content = grid;
            root.Children.Add(scrollViewer);

            var buttonsPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 14, 0, 0)
            };
            Grid.SetRow(buttonsPanel, 2);

            var okButton = CreateActionButton("ОК");
            okButton.Margin = new Thickness(0, 0, 10, 0);
            var cancelButton = CreateSecondaryButton("Отмена");
            cancelButton.Margin = new Thickness(0);

            bool merged = false;

            okButton.Click += (okSender, args) =>
            {
                var selectedCells = buttonMap
                    .Where(pair => pair.Key.IsChecked == true)
                    .Select(pair => pair.Value)
                    .ToList();

                if (selectedCells.Count < 2)
                {
                    MessageBox.Show("Выберите минимум две соседние ячейки.", "Таблица",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (!TryMergeTableCells(structure, selectedCells, out string errorMessage))
                {
                    MessageBox.Show(errorMessage, "Таблица",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                merged = true;
                window.DialogResult = true;
                window.Close();
            };

            cancelButton.Click += (cancelSender, args) =>
            {
                window.DialogResult = false;
                window.Close();
            };

            buttonsPanel.Children.Add(okButton);
            buttonsPanel.Children.Add(cancelButton);
            root.Children.Add(buttonsPanel);

            window.Content = root;
            window.ShowDialog();

            return merged;
        }

        private int GetDisplayRow(TableStructure structure, TableCellDefinition cell)
        {
            return structure?.HeaderCells.Contains(cell) == true
                ? cell.Row
                : (structure?.HeaderRowCount ?? 0) + cell.Row;
        }

        private bool TryMergeTableCells(TableStructure structure, List<TableCellDefinition> selectedCells, out string errorMessage)
        {
            errorMessage = null;

            if (selectedCells.Any(cell => cell.ColSpan > 1 || cell.RowSpan > 1))
            {
                errorMessage = "Нельзя повторно объединять уже объединенные ячейки.";
                return false;
            }

            bool isHeaderSelection = structure.HeaderCells.Contains(selectedCells[0]);
            bool isBodySelection = structure.BodyCells.Contains(selectedCells[0]);

            if (selectedCells.Any(cell =>
                    structure.HeaderCells.Contains(cell) != isHeaderSelection ||
                    structure.BodyCells.Contains(cell) != isBodySelection))
            {
                errorMessage = "Нельзя объединять вместе ячейки шапки и строки данных.";
                return false;
            }

            int minRow = selectedCells.Min(cell => cell.Row);
            int maxRow = selectedCells.Max(cell => cell.Row);
            int minColumn = selectedCells.Min(cell => cell.Column);
            int maxColumn = selectedCells.Max(cell => cell.Column);
            int expectedCount = (maxRow - minRow + 1) * (maxColumn - minColumn + 1);

            if (selectedCells.Count != expectedCount)
            {
                errorMessage = "Ячейки должны образовывать сплошной соседний прямоугольник.";
                return false;
            }

            for (int row = minRow; row <= maxRow; row++)
            {
                for (int column = minColumn; column <= maxColumn; column++)
                {
                    if (!selectedCells.Any(cell => cell.Row == row && cell.Column == column))
                    {
                        errorMessage = "Ячейки должны быть соседними.";
                        return false;
                    }
                }
            }

            var topLeftCell = selectedCells
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .First();

            topLeftCell.ColSpan = maxColumn - minColumn + 1;
            topLeftCell.RowSpan = maxRow - minRow + 1;

            var targetCollection = isHeaderSelection ? structure.HeaderCells : structure.BodyCells;

            foreach (var cell in selectedCells.Where(cell => !ReferenceEquals(cell, topLeftCell)).ToList())
            {
                targetCollection.Remove(cell);
            }

            NormalizeTableStructure(structure);
            return true;
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

        private void MenuItem_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border)
            {
                ActivateMenuItem(border);
            }
        }

        private void ActivateMenuItem(Border border, bool showContent = true, bool editMode = false)
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
                ShowContentForMenuItem(border, editMode);
            }
            else if (dynamicTemplates.ContainsKey(border)
                || dynamicSections.ContainsKey(border)
                || dynamicSubSections.ContainsKey(border))
            {
                MainContentControl.Content = CreateDefaultContent();
                currentDynamicEditorBorder = null;
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

        private void ShowContentForMenuItem(Border menuItem, bool editMode = false)
        {
            if (TryShowDynamicContent(menuItem, editMode))
            {
                UpdateHistoryCalendarVisibility(false);
                return;
            }

            currentDynamicEditorBorder = null;
            UpdateHistoryCalendarVisibility(false);

            if (menuItem == Item_Archive)
            {
                MainContentControl.Content = new ArchivePage(
                    GetArchivedReportsFolderPath(),
                    DownloadArchivedReportAsync,
                    DeleteArchivedReportAsync);
            }
            else if (menuItem == Item_Templates)
            {
                MainContentControl.Content = new TemplatesPage(
                    GetSavedTemplatesFolderPath(),
                    RestoreSavedTemplateSnapshotAsync,
                    DeleteSavedTemplateSnapshotAsync);
            }
            else
            {
                MainContentControl.Content = CreateDefaultContent();
                currentDynamicEditorBorder = null;
            }
        }

        private void Item_Archive_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
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
            FilterMenuItem(Item_Archive, query);
            FilterMenuItem(Item_Templates, query);
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
