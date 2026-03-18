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
using Project_bpi.Models;
using Project_bpi.Services;

namespace Project_bpi
{
    public partial class MainWindow : Window
    {
        public DataBase DB = new DataBase();
        private const string TemplateDatabaseFolderName = "TemplateDatabases";
        private const string SectionContentSubSectionTitle = "__section_content__";
        private const string StudyReportTitle = "Учебный отчет";
        private const string NirReportTitle = "Отчет по НИР";
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
        }

        public MainWindow()
        {
            DB.InitializeDatabase();
            InitializeComponent();
            InitializeDateRange();
            InitializeMenuHierarchy();
            HideStaticNirReportMenu();
            HideStaticStudyReportMenu();
            Loaded += async (sender, args) =>
            {
                await EnsureNirReportInDatabaseAsync();
                await EnsureStudyReportInDatabaseAsync();
                await LoadSavedTemplatesAsync();
            };
        }

        private void InitializeMenuHierarchy()
        {
            // 1-й уровень → корневой заголовок НИР
            parents[NIR_Header] = null; // корень
            parents[Section1_Header] = NIR_Header;
            parents[Item_Title] = NIR_Header;
            parents[Section2_Header] = NIR_Header;
            parents[Section3_Header] = NIR_Header;
            parents[Section4_Header] = NIR_Header;
            parents[Section5_Header] = NIR_Header;
            parents[Section6_Header] = NIR_Header;
            parents[Section7_Header] = NIR_Header;
            parents[Section8_Header] = NIR_Header;
            parents[Section9_Header] = NIR_Header;

            // 2-й уровень → Раздел 1
            parents[Item_11] = Section1_Header;
            parents[Item_12] = Section1_Header;
            parents[Item_13] = Section1_Header;
            parents[Item_14] = Section1_Header;
            parents[Item_15] = Section1_Header;

            // 2-й уровень → Раздел 2
            parents[Item_21] = Section2_Header;
            parents[Item_22] = Section2_Header;
            parents[Item_23] = Section2_Header;
            parents[Item_24] = Section2_Header;
            parents[Item_25] = Section2_Header;

            // 2-й уровень → Раздел 3
            parents[Item_31] = Section3_Header;
            parents[Item_32] = Section3_Header;

            // 2-й уровень → Раздел 4
            parents[Item_41] = Section4_Header;
            parents[Item_42] = Section4_Header;
            parents[Item_43] = Section4_Header;

            // 2-й уровень → Раздел 5
            parents[Item_51] = Section5_Header;
            parents[Item_52] = Section5_Header;
            parents[Item_53] = Section5_Header;
            parents[Item_54] = Section5_Header;

            // 3-й уровень → Подразделы внутри "Реализуемые стартап-проекты"
            parents[Item_51a] = Item_51;

            // 3-й уровень → Подразделы внутри "Научные конференции"
            parents[Item_53a] = Item_53;
            parents[Item_53b] = Item_53;
            // Учебный отчет → корень
            parents[Study_Header] = null;

            // Раздел 14 → внутри Учебного отчета
            parents[Section14_Header] = Study_Header;

            // Подразделы 14.x
            parents[Item_141] = Section14_Header;
            parents[Item_142] = Section14_Header;
            parents[Item_143] = Section14_Header;
        }

        private void HideStaticStudyReportMenu()
        {
            Study_Header.Visibility = Visibility.Collapsed;
            Study_Menu.Visibility = Visibility.Collapsed;
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

            DynamicTemplatesPanel.Children.Add(templateContainer);
            return entry;
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
            var actionsPanel = CreateDynamicActionsPanel(editButton, deleteButton);

            border.MouseEnter += (sender, args) =>
            {
                editButton.Visibility = Visibility.Visible;
                deleteButton.Visibility = Visibility.Visible;
            };
            border.MouseLeave += (sender, args) =>
            {
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
                    MainContentControl.Content = null;
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
                    stack.Children.Add(CreateSubSectionPreviewTableCard(table));
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
            string sectionTitle = subsection.Section != null
                ? GetSectionContentTitle(subsection.Section)
                : string.Empty;
            string subSectionTitle = BuildSubSectionTitle(subsection.Section, subsection);
            string content = subsection.Texts != null && subsection.Texts.Any()
                ? subsection.Texts.First().Content
                : "Текст подраздела пока не заполнен.";

            var stack = new StackPanel
            {
                Margin = new Thickness(20)
            };

            stack.Children.Add(CreateSubSectionPreviewHeader(sectionTitle, subSectionTitle));
            stack.Children.Add(CreateSubSectionPreviewBody(content));

            if (subsection.Tables != null && subsection.Tables.Any())
            {
                foreach (var table in subsection.Tables.OrderBy(t => t.Id))
                {
                    stack.Children.Add(CreateSubSectionPreviewTableCard(table));
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
            var content = new StackPanel();

            content.Children.Add(new TextBlock
            {
                Text = table.Title,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 8),
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4")),
                TextWrapping = TextWrapping.Wrap
            });

            content.Children.Add(CreateTablePreview(table));

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
                TableEditorHost = editorHost
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
            panel.Children.Add(tableActions);

            panel.Children.Add(new TextBlock
            {
                Text = "Нажмите на ячейку и печатайте прямо в таблице. Первая строка относится к шапке, под ней автоматически показывается нумерация столбцов, новые строки добавляются как строки данных.",
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
            context.TableEditorHost.Content = CreateEditableTableGrid(context);
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
            bool hasAutoNumberRow = ShouldShowAutoNumberRow(structure);
            int numberRowOffset = hasAutoNumberRow ? 1 : 0;
            int totalDisplayRowCount = structure.HeaderRowCount + numberRowOffset + structure.BodyRowCount;

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
            if (hasAutoNumberRow)
            {
                AddAutoNumberRowToGrid(grid, structure.HeaderRowCount + 1, structure.ColumnCount, 1);
            }

            AddEditableCellsToGrid(grid, structure.BodyCells, structure.HeaderRowCount + numberRowOffset + 1, 1);

            return new Border
            {
                BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#ced4da")),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Child = grid
            };
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

        private void AddColumnHandles(Grid grid, TableEditorContext context)
        {
            for (int column = 1; column <= context.Structure.ColumnCount; column++)
            {
                int targetColumn = column;
                bool canDelete = context.Structure.ColumnCount > 1;
                var handle = CreateTableEdgeHandle(
                    plusToolTip: "Вставить столбец",
                    minusToolTip: "Удалить столбец",
                    onPlusClick: () =>
                    {
                        InsertTableColumn(context, targetColumn);
                    },
                    onMinusClick: canDelete
                        ? (System.Action)(() => DeleteTableColumn(context, targetColumn))
                        : null,
                    isColumnHandle: true);

                Grid.SetRow(handle, 0);
                Grid.SetColumn(handle, targetColumn);
                grid.Children.Add(handle);
            }

            var trailingHandle = CreateTableEdgeHandle(
                plusToolTip: "Добавить столбец справа",
                minusToolTip: null,
                onPlusClick: () =>
                {
                    InsertTableColumn(context, context.Structure.ColumnCount + 1);
                },
                onMinusClick: null,
                isColumnHandle: true);

            Grid.SetRow(trailingHandle, 0);
            Grid.SetColumn(trailingHandle, context.Structure.ColumnCount + 1);
            grid.Children.Add(trailingHandle);
        }

        private void AddRowHandles(Grid grid, TableEditorContext context)
        {
            bool hasAutoNumberRow = ShouldShowAutoNumberRow(context.Structure);
            int numberRowOffset = hasAutoNumberRow ? 1 : 0;
            int totalStructureRowCount = context.Structure.HeaderRowCount + context.Structure.BodyRowCount;

            for (int structureRow = 1; structureRow <= totalStructureRowCount; structureRow++)
            {
                bool isHeaderRow = structureRow <= context.Structure.HeaderRowCount;
                bool canDelete = isHeaderRow
                    ? context.Structure.HeaderRowCount > 1
                    : context.Structure.BodyRowCount > 0;
                int gridRow = isHeaderRow
                    ? structureRow
                    : structureRow + numberRowOffset;

                var handle = CreateTableEdgeHandle(
                    plusToolTip: "Вставить строку",
                    minusToolTip: "Удалить строку",
                    onPlusClick: () =>
                    {
                        InsertTableRow(context, structureRow);
                    },
                    onMinusClick: canDelete
                        ? (System.Action)(() => DeleteTableRow(context, structureRow))
                        : null,
                    isColumnHandle: false);

                Grid.SetRow(handle, gridRow);
                Grid.SetColumn(handle, 0);
                grid.Children.Add(handle);
            }

            var trailingHandle = CreateTableEdgeHandle(
                plusToolTip: "Добавить строку снизу",
                minusToolTip: null,
                onPlusClick: () =>
                {
                    InsertTableRow(context, totalStructureRowCount + 1);
                },
                onMinusClick: null,
                isColumnHandle: false);

            Grid.SetRow(trailingHandle, totalStructureRowCount + numberRowOffset + 1);
            Grid.SetColumn(trailingHandle, 0);
            grid.Children.Add(trailingHandle);
        }

        private Border CreateTableEdgeHandle(string plusToolTip, string minusToolTip, System.Action onPlusClick, System.Action onMinusClick, bool isColumnHandle)
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

        private TextBox CreateEditableTableCellTextBox(TableCellDefinition cell)
        {
            var textBox = new TextBox
            {
                Text = cell.Text ?? string.Empty,
                MinWidth = 140,
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
                VerticalContentAlignment = VerticalAlignment.Center,
                HorizontalContentAlignment = cell.IsHeader ? HorizontalAlignment.Center : HorizontalAlignment.Left,
                TextAlignment = cell.IsHeader ? TextAlignment.Center : TextAlignment.Left
            };

            textBox.TextChanged += (sender, args) =>
            {
                cell.Text = textBox.Text;
            };

            return textBox;
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

            var database = new DataBase(templateEntry.DatabasePath);
            templateEntry.Report.Title = newTitle;
            templateEntry.Report.PatternFile.Title = newTitle;

            await database.UpdateReport(templateEntry.Report);
            await database.UpdateFilePattern(templateEntry.Report.PatternFile);

            templateEntry.DisplayTitle = newTitle;
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

            await SaveTableAsync(context, true);
        }

        private async Task<bool> SaveTableAsync(TableEditorContext context, bool refreshAfterSave)
        {
            if (context == null)
            {
                return false;
            }

            string title = context.TitleTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(title))
            {
                MessageBox.Show("Название таблицы не может быть пустым.", "Таблица",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

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
                    IsHeader = true
                });
            }

            foreach (var bodyCell in structure.BodyCells)
            {
                await database.AddTableItem(new TableItem
                {
                    TableId = context.Table.Id,
                    Row = bodyCell.Row,
                    Column = bodyCell.Column,
                    Header = bodyCell.Text,
                    ColSpan = bodyCell.ColSpan,
                    RowSpan = bodyCell.RowSpan,
                    IsHeader = false
                });
            }

            if (refreshAfterSave)
            {
                await RefreshTemplateEntryAsync(context.Template, subsectionId: context.SubSection.Id);
            }

            return true;
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
            bool hasAutoNumberRow = ShouldShowAutoNumberRow(structure);

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
                for (int index = 0; index < columnCount; index++)
                {
                    headerGrid.ColumnDefinitions.Add(new ColumnDefinition
                    {
                        Width = index == 0 ? new GridLength(50) : new GridLength(1, GridUnitType.Star)
                    });
                }

                for (int rowIndex = 0; rowIndex < structure.HeaderRowCount; rowIndex++)
                {
                    headerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                }

                AddHeaderCells(headerGrid, structure, columnCount);
                panel.Children.Add(headerGrid);
            }

            if (hasAutoNumberRow)
            {
                panel.Children.Add(CreateAutoNumberRowGrid(columnCount));
            }

            if (structure.BodyRowCount > 0)
            {
                var bodyGrid = new Grid();
                for (int index = 0; index < columnCount; index++)
                {
                    bodyGrid.ColumnDefinitions.Add(new ColumnDefinition
                    {
                        Width = index == 0 ? new GridLength(50) : new GridLength(1, GridUnitType.Star)
                    });
                }

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
                Child = panel
            };
        }

        private bool ShouldShowAutoNumberRow(TableStructure structure)
        {
            return structure != null
                && structure.HeaderRowCount > 0
                && structure.ColumnCount > 0;
        }

        private Grid CreateAutoNumberRowGrid(int columnCount)
        {
            var grid = new Grid
            {
                Background = Brushes.White
            };

            for (int index = 0; index < columnCount; index++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition
                {
                    Width = index == 0 ? new GridLength(50) : new GridLength(1, GridUnitType.Star)
                });
            }

            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            AddAutoNumberRowToGrid(grid, 0, columnCount, 0);
            return grid;
        }

        private void AddAutoNumberRowToGrid(Grid grid, int gridRowIndex, int columnCount, int columnOffset)
        {
            for (int column = 1; column <= columnCount; column++)
            {
                var numberBorder = CreateAutoNumberCellBorder(column.ToString());
                Grid.SetRow(numberBorder, gridRowIndex);
                Grid.SetColumn(numberBorder, columnOffset + column - 1);
                grid.Children.Add(numberBorder);
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
                var border = CreateTableCellBorder(cell.Text, isHeader);
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
                Background = Brushes.White
            };

            for (int index = 0; index < columnCount; index++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition
                {
                    Width = index == 0 ? new GridLength(50) : new GridLength(1, GridUnitType.Star)
                });
            }

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
                Child = new TextBlock
                {
                    Text = text,
                    FontSize = isHeader ? 10 : 12,
                    FontWeight = isHeader ? FontWeights.Bold : FontWeights.Normal,
                    Foreground = Brushes.Black,
                    TextAlignment = TextAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextWrapping = TextWrapping.Wrap
                }
            };
        }

        private Border CreateAutoNumberCellBorder(string text)
        {
            return new Border
            {
                BorderBrush = Brushes.Black,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Padding = new Thickness(8, 6, 8, 6),
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#f8f9fa")),
                Child = new TextBlock
                {
                    Text = text,
                    FontSize = 10,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.Black,
                    TextAlignment = TextAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextWrapping = TextWrapping.NoWrap
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
                .OrderBy(c => GetDisplayRow(c, structure.HeaderRowCount))
                .ThenBy(c => c.Column))
            {
                int displayRow = GetDisplayRow(cell, structure.HeaderRowCount);
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

        private int GetDisplayRow(TableCellDefinition cell, int headerRowCount)
        {
            return cell.IsHeader ? cell.Row : headerRowCount + cell.Row;
        }

        private bool TryMergeTableCells(TableStructure structure, List<TableCellDefinition> selectedCells, out string errorMessage)
        {
            errorMessage = null;

            if (selectedCells.Any(cell => cell.ColSpan > 1 || cell.RowSpan > 1))
            {
                errorMessage = "Нельзя повторно объединять уже объединенные ячейки.";
                return false;
            }

            bool isHeaderSelection = selectedCells[0].IsHeader;
            if (selectedCells.Any(cell => cell.IsHeader != isHeaderSelection))
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

            var targetCollection = isHeaderSelection
                ? structure.HeaderCells
                : structure.BodyCells;

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

        // === МЕНЮ ===

        private void NIR_Header_Click(object sender, MouseButtonEventArgs e)
        {
            ToggleSectionMenu(NIR_Menu, NIR_Arrow);
            MenuItem_Click(sender, e);
        }

        private void Section1_Header_Click(object sender, MouseButtonEventArgs e)
        {
            ToggleSectionMenu(Section1_Menu, NIR_Arrow1);
            MenuItem_Click(sender, e);
        }

        private void Section2_Header_Click(object sender, MouseButtonEventArgs e)
        {
            ToggleSectionMenu(Section2_Menu, NIR_Arrow2);
            MenuItem_Click(sender, e);
        }

        private void Section3_Header_Click(object sender, MouseButtonEventArgs e)
        {
            ToggleSectionMenu(Section3_Menu, NIR_Arrow3);
            MenuItem_Click(sender, e);
        }

        private void Section4_Header_Click(object sender, MouseButtonEventArgs e)
        {
            ToggleSectionMenu(Section4_Menu, NIR_Arrow4);
            MenuItem_Click(sender, e);
        }

        private void Section5_Header_Click(object sender, MouseButtonEventArgs e)
        {
            ToggleSectionMenu(Section5_Menu, NIR_Arrow5);
            MenuItem_Click(sender, e);
        }

        // Простые разделы без подменю
        private void Section7_Header_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Section8_Header_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Section9_Header_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Study_Header_Click(object sender, MouseButtonEventArgs e)
        {
            ToggleSectionMenu(Study_Menu, Study_Arrow);
            MenuItem_Click(sender, e);
        }

        private void Section14_Header_Click(object sender, MouseButtonEventArgs e)
        {
            ToggleSectionMenu(Section14_Menu, Study_Arrow14);
            MenuItem_Click(sender, e);
        }

        // Обработчики подразделов
        private void Item_141_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_142_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_143_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void ToggleSectionMenu(StackPanel menu, Image arrow)
        {
            if (menu.Visibility == Visibility.Visible)
            {
                menu.Visibility = Visibility.Collapsed;
                ApplyIndicatorVisual(arrow, false, false);
            }
            else
            {
                menu.Visibility = Visibility.Visible;
                ApplyIndicatorVisual(arrow, true, false);
            }
        }

        // Вложенные подразделы
        private void Item_51_Click(object sender, MouseButtonEventArgs e)
        {
            ToggleSectionMenu(Item_51_SubMenu, NIR_Arrow51);
            MenuItem_Click(sender, e);
        }

        private void Item_53_Click(object sender, MouseButtonEventArgs e)
        {
            ToggleSectionMenu(Item_53_SubMenu, NIR_Arrow53);
            MenuItem_Click(sender, e);
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
                return;
            }

            currentDynamicEditorBorder = null;

            if (menuItem == Item_11)
            {
                MainContentControl.Content = new Item11View();
            }
            else if (menuItem == Item_12)
            {
                MainContentControl.Content = new ContractResearchView();
            }
            else if (menuItem == Item_13)
            {
                MainContentControl.Content = new Item_13View();
            }
            else if (menuItem == Item_14)
            {
                MainContentControl.Content = new Item_14View();
            }
            else if (menuItem == Item_15)
            {
                MainContentControl.Content = new СooperationProductionView();
            }
            else if (menuItem == Item_21)
            {
                MainContentControl.Content = new Item21_View();
            }
            else if (menuItem == Item_22)
            {
                MainContentControl.Content = new Item22_View();
            }
            else if (menuItem == Item_23)
            {
                MainContentControl.Content = new Item23_View();
            }
            else if (menuItem == Item_24)
            {
                MainContentControl.Content = new Item24_View();
            }
            else if (menuItem == Item_25)
            {
                MainContentControl.Content = new Item25_View();
            }
            else if (menuItem == Item_31)
            {
                MainContentControl.Content = new Item31_View();
            }
            else if (menuItem == Item_32)
            {
                MainContentControl.Content = new Item32_View();
            }
            else if (menuItem == Section4_Header)
            {
                MainContentControl.Content = new Item4_View();
            }
            else if (menuItem == Item_41)
            {
                MainContentControl.Content = new Item41_View();
            }
            else if (menuItem == Item_42)
            {
                MainContentControl.Content = new Item42_View();
            }
            else if (menuItem == Item_43)
            {
                MainContentControl.Content = new Item43_View();
            }
            else if (menuItem == Item_51)
            {
                MainContentControl.Content = new Item51_View();
            }
            else if (menuItem == Item_51a)
            {
                MainContentControl.Content = new Item51a_View();
            }
            else if (menuItem == Item_52)
            {
                MainContentControl.Content = new Item52_View();
            }
            else if (menuItem == Item_53)
            {
                MainContentControl.Content = new Item53_View();
            }
            else if (menuItem == Item_53a)
            {
                MainContentControl.Content = new Item53a_View();
            }
            else if (menuItem == Item_53b)
            {
                MainContentControl.Content = new Item53b_View();
            }
            else if (menuItem == Item_54)
            {
                MainContentControl.Content = new Item54_View();
            }
            else if (menuItem == Section6_Header)
            {
                MainContentControl.Content = new Item6_View();
            }
            else if (menuItem == Section7_Header)
            {
                MainContentControl.Content = new Item7_View();
            }
            else if (menuItem == Section8_Header)
            {
                MainContentControl.Content = new Item8_View();
            }
            else if (menuItem == Item_Archive)
            {
                MainContentControl.Content = new ArchivePage();
            }
            else if (menuItem == Item_Templates)
            {
                MainContentControl.Content = new TemplatesPage();
            }
            else if (menuItem == Item_141)
            {
                MainContentControl.Content = new Item141_View();
            }
            else if (menuItem == Item_142)
            {
                MainContentControl.Content = new Item142_View();
            }
            else if (menuItem == Item_143)
            {
                MainContentControl.Content = new Item143_View();
            }
            else if (menuItem == Section7_Header || menuItem == Section8_Header || menuItem == Section9_Header)
            {
                MainContentControl.Content = new TextBlock
                {
                    Text = $"Содержимое {menuItem.Name}",
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Foreground = Brushes.Gray,
                    FontSize = 16
                };
            }
            else
            {
                MainContentControl.Content = CreateDefaultContent();
                currentDynamicEditorBorder = null;
            }
        }

        // Обработчики кликов по новым пунктам
        private void Item_21_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_22_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_23_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_24_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_25_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);

        private void Item_31_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_32_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);

        private void Item_41_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_42_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_43_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);

        private void Item_51a_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_52_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_53a_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_53b_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_54_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);

        private void Item_11_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_12_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_13_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_14_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_15_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void Item_Archive_Click(object sender, MouseButtonEventArgs e) => MenuItem_Click(sender, e);
        private void HistoryButton_Click(object sender, RoutedEventArgs e)
        {
            MainContentControl.Content = new HistoryChangesView();
        }
        private void HelpHyperlink_Click(object sender, RoutedEventArgs e)
        {
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
            ApplyIndicatorVisual(NIR_Arrow, NIR_Menu.Visibility == Visibility.Visible, IsMenuItemActive(NIR_Header));
            ApplyIndicatorVisual(NIR_Arrow1, Section1_Menu.Visibility == Visibility.Visible, IsMenuItemActive(Section1_Header));
            ApplyIndicatorVisual(NIR_Arrow2, Section2_Menu.Visibility == Visibility.Visible, IsMenuItemActive(Section2_Header));
            ApplyIndicatorVisual(NIR_Arrow3, Section3_Menu.Visibility == Visibility.Visible, IsMenuItemActive(Section3_Header));
            ApplyIndicatorVisual(NIR_Arrow4, Section4_Menu.Visibility == Visibility.Visible, IsMenuItemActive(Section4_Header));
            ApplyIndicatorVisual(NIR_Arrow5, Section5_Menu.Visibility == Visibility.Visible, IsMenuItemActive(Section5_Header));
            ApplyIndicatorVisual(NIR_Arrow51, Item_51_SubMenu.Visibility == Visibility.Visible, IsMenuItemActive(Item_51));
            ApplyIndicatorVisual(NIR_Arrow53, Item_53_SubMenu.Visibility == Visibility.Visible, IsMenuItemActive(Item_53));
            ApplyIndicatorVisual(Study_Arrow, Study_Menu.Visibility == Visibility.Visible, IsMenuItemActive(Study_Header));
            ApplyIndicatorVisual(Study_Arrow14, Section14_Menu.Visibility == Visibility.Visible, IsMenuItemActive(Section14_Header));

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

            FilterMenuPanel(NIR_Menu, query);
            FilterMenuPanel(Study_Menu, query);
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

                    StackPanel subPanel = null;
                    if (border.Name == "Section1_Header") subPanel = Section1_Menu;
                    else if (border.Name == "Section2_Header") subPanel = Section2_Menu;
                    else if (border.Name == "Section3_Header") subPanel = Section3_Menu;
                    else if (border.Name == "Section4_Header") subPanel = Section4_Menu;
                    else if (border.Name == "Section5_Header") subPanel = Section5_Menu;
                    else if (border.Name == "Item_51") subPanel = Item_51_SubMenu;
                    else if (border.Name == "Item_53") subPanel = Item_53_SubMenu;
                    else if (border.Name == "NIR_Header") subPanel = NIR_Menu;
                    else if (border.Name == "Section14_Header") subPanel = Section14_Menu;
                    else if (border.Name == "Study_Header") subPanel = Study_Menu;

                    if (subPanel != null)
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
