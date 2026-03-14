using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Globalization;
using Project_bpi.Models;
using Project_bpi.Services;

namespace Project_bpi
{
    public partial class MainWindow : Window
    {
        public DataBase DB = new DataBase();
        private readonly DataBase TemplateDb = new DataBase("test.db");

        private readonly Dictionary<Border, Border> parents = new Dictionary<Border, Border>();
        private readonly Dictionary<Border, DynamicTemplateEntry> dynamicTemplates = new Dictionary<Border, DynamicTemplateEntry>();
        private readonly Dictionary<Border, Section> dynamicSections = new Dictionary<Border, Section>();
        private readonly Dictionary<Border, SubSection> dynamicSubSections = new Dictionary<Border, SubSection>();
        private readonly Dictionary<Border, StackPanel> dynamicMenus = new Dictionary<Border, StackPanel>();
        private readonly Dictionary<Border, TextBlock> dynamicIndicators = new Dictionary<Border, TextBlock>();
        private DateTime _currentCalendarDate = DateTime.Today;

        private sealed class DynamicTemplateEntry
        {
            public string DisplayTitle { get; set; }
            public Report Report { get; set; }
        }

        public MainWindow()
        {
            DB.InitializeDatabase();
            InitializeComponent();
            InitializeDateRange();
            InitializeMenuHierarchy();
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

        private async void AddTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new TemplateNameDialog
            {
                Owner = this,
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
                MessageBox.Show("Шаблон с таким названием уже добавлен в меню.", "Шаблоны",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                TemplateDb.InitializeDatabase();

                var reports = await TemplateDb.GetAllReports();
                if (reports.Count == 0)
                {
                    MessageBox.Show("В test.db не найдено ни одного отчета для построения шаблона.", "Шаблоны",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                Report report = await TemplateDb.GetFullReport(reports[0].Id);
                if (report == null)
                {
                    MessageBox.Show("Не удалось прочитать структуру шаблона из test.db.", "Шаблоны",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                AddTemplateToMenu(templateTitle, report);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось добавить шаблон из test.db:{Environment.NewLine}{ex.Message}",
                    "Шаблоны", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddTemplateToMenu(string templateTitle, Report report)
        {
            var templateContainer = new StackPanel();
            var templateMenu = new StackPanel { Visibility = Visibility.Collapsed };
            var templateHeader = CreateDynamicMenuBorder(templateTitle, "main", "MenuHeaderStyle", true, out var templateIndicator);

            dynamicTemplates[templateHeader] = new DynamicTemplateEntry
            {
                DisplayTitle = templateTitle,
                Report = report
            };
            dynamicMenus[templateHeader] = templateMenu;
            dynamicIndicators[templateHeader] = templateIndicator;
            parents[templateHeader] = null;

            templateHeader.MouseLeftButtonDown += DynamicTemplateHeader_Click;

            templateContainer.Children.Add(templateHeader);
            templateContainer.Children.Add(templateMenu);

            foreach (var section in report.Sections.OrderBy(s => s.Number))
            {
                section.Report = report;

                var sectionContainer = new StackPanel();
                bool hasChildren = section.SubSections != null && section.SubSections.Any();
                var sectionHeader = CreateDynamicMenuBorder(BuildSectionTitle(section), "sub", "SubMenuStyle", hasChildren, out var sectionIndicator);

                dynamicSections[sectionHeader] = section;
                parents[sectionHeader] = templateHeader;
                sectionHeader.MouseLeftButtonDown += DynamicSectionHeader_Click;

                sectionContainer.Children.Add(sectionHeader);

                if (hasChildren)
                {
                    var sectionMenu = new StackPanel { Visibility = Visibility.Collapsed };
                    dynamicMenus[sectionHeader] = sectionMenu;
                    dynamicIndicators[sectionHeader] = sectionIndicator;

                    foreach (var subsection in section.SubSections.OrderBy(s => s.Number))
                    {
                        subsection.Section = section;

                        var subsectionHeader = CreateDynamicMenuBorder(
                            BuildSubSectionTitle(section, subsection),
                            "sub2",
                            "SubMenuStyle1",
                            false,
                            out _);

                        dynamicSubSections[subsectionHeader] = subsection;
                        parents[subsectionHeader] = sectionHeader;
                        subsectionHeader.MouseLeftButtonDown += DynamicLeaf_Click;

                        sectionMenu.Children.Add(subsectionHeader);
                    }

                    sectionContainer.Children.Add(sectionMenu);
                }

                templateMenu.Children.Add(sectionContainer);
            }

            DynamicTemplatesPanel.Children.Add(templateContainer);
        }

        private Border CreateDynamicMenuBorder(string text, string tag, string textStyleKey, bool expandable, out TextBlock indicator)
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

            indicator = new TextBlock
            {
                Text = ">",
                Margin = new Thickness(6, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4")),
                FontWeight = FontWeights.Bold
            };

            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal
            };

            panel.Children.Add(title);
            panel.Children.Add(indicator);
            border.Child = panel;

            return border;
        }

        private void DynamicTemplateHeader_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border)
            {
                ToggleDynamicMenu(border);
                MenuItem_Click(sender, e);
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

                MenuItem_Click(sender, e);
            }
        }

        private void DynamicLeaf_Click(object sender, MouseButtonEventArgs e)
        {
            MenuItem_Click(sender, e);
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
                indicator.Text = isExpanded ? "v" : ">";
            }
        }

        private string BuildSectionTitle(Section section)
        {
            if (string.IsNullOrWhiteSpace(section.Title))
            {
                return $"Раздел {section.Number}";
            }

            return $"Раздел {section.Number}. {section.Title}";
        }

        private string BuildSubSectionTitle(Section section, SubSection subsection)
        {
            string prefix = section != null
                ? $"{section.Number}.{subsection.Number}"
                : subsection.Number.ToString();

            if (string.IsNullOrWhiteSpace(subsection.Title))
            {
                return prefix;
            }

            return $"{prefix} {subsection.Title}";
        }

        private bool TryShowDynamicContent(Border menuItem)
        {
            if (dynamicTemplates.TryGetValue(menuItem, out var templateEntry))
            {
                MainContentControl.Content = CreateReportContent(templateEntry);
                return true;
            }

            if (dynamicSections.TryGetValue(menuItem, out var section))
            {
                MainContentControl.Content = CreateSectionContent(section);
                return true;
            }

            if (dynamicSubSections.TryGetValue(menuItem, out var subsection))
            {
                MainContentControl.Content = CreateSubSectionContent(subsection);
                return true;
            }

            return false;
        }

        private UIElement CreateReportContent(DynamicTemplateEntry templateEntry)
        {
            var stack = CreateContentStack(templateEntry.DisplayTitle,
                $"Структура взята из test.db. Отчет-источник: {templateEntry.Report.Title}");

            foreach (var section in templateEntry.Report.Sections.OrderBy(s => s.Number))
            {
                var sectionLines = section.SubSections != null && section.SubSections.Any()
                    ? section.SubSections.OrderBy(s => s.Number).Select(s => BuildSubSectionTitle(section, s))
                    : new[] { "Подразделы отсутствуют" };

                stack.Children.Add(CreateTextCard(BuildSectionTitle(section), string.Join(Environment.NewLine, sectionLines)));
            }

            return new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = stack
            };
        }

        private UIElement CreateSectionContent(Section section)
        {
            var stack = CreateContentStack(BuildSectionTitle(section),
                "Раздел загружен из test.db и готов для дальнейшей привязки к экрану.");

            var subsectionLines = section.SubSections != null && section.SubSections.Any()
                ? section.SubSections.OrderBy(s => s.Number).Select(s => BuildSubSectionTitle(section, s))
                : new[] { "Подразделы отсутствуют" };

            stack.Children.Add(CreateTextCard("Подразделы", string.Join(Environment.NewLine, subsectionLines)));

            return new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = stack
            };
        }

        private UIElement CreateSubSectionContent(SubSection subsection)
        {
            string title = BuildSubSectionTitle(subsection.Section, subsection);
            var stack = CreateContentStack(title,
                "Подраздел загружен из test.db. Ниже показаны текстовые блоки и таблицы, если они есть.");

            if (subsection.Texts != null && subsection.Texts.Any())
            {
                foreach (var text in subsection.Texts)
                {
                    stack.Children.Add(CreateTextCard($"Текст ({text.PatternName})", text.Content));
                }
            }
            else
            {
                stack.Children.Add(CreateTextCard("Текст", "Для этого подраздела текстовые блоки в базе не найдены."));
            }

            if (subsection.Tables != null && subsection.Tables.Any())
            {
                foreach (var table in subsection.Tables)
                {
                    stack.Children.Add(CreateContentCard(table.Title, CreateTablePreview(table)));
                }
            }

            return new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = stack
            };
        }

        private StackPanel CreateContentStack(string title, string subtitle)
        {
            var stack = new StackPanel
            {
                Margin = new Thickness(24)
            };

            stack.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 22,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#0167a4")),
                TextWrapping = TextWrapping.Wrap
            });

            stack.Children.Add(new TextBlock
            {
                Text = subtitle,
                Margin = new Thickness(0, 8, 0, 18),
                Foreground = Brushes.DimGray,
                TextWrapping = TextWrapping.Wrap
            });

            return stack;
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
            var panel = new StackPanel();
            var groupedRows = table.TableItems
                .OrderBy(item => item.Row)
                .ThenBy(item => item.Column)
                .GroupBy(item => item.Row);

            foreach (var row in groupedRows)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = string.Join(" | ", row.Select(item => item.Header)),
                    Margin = new Thickness(0, 0, 0, 4),
                    TextWrapping = TextWrapping.Wrap,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#343a40"))
                });
            }

            return panel;
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
                arrow.LayoutTransform = null;
            }
            else
            {
                menu.Visibility = Visibility.Visible;
                arrow.LayoutTransform = new RotateTransform(90);
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
            ResetAllStyles();

            Border b = sender as Border;
            if (b == null) return;

            HighlightChain(b);
            currentActive = b;
            ShowContentForMenuItem(b);
            UpdateArrows();
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
                return;
            }

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
                MainContentControl.Content = new TextBlock
                {
                    Text = "Основная область контента",
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Foreground = Brushes.Gray,
                    FontSize = 16
                };
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

        private bool IsActive(Border b)
        {
            return b.Style == (Style)FindResource("ActiveMainItemStyle")
                || b.Style == (Style)FindResource("ActiveSubItemStyle")
                || b.Style == (Style)FindResource("ActiveSubItemLevel2Style");
        }

        private void UpdateArrows()
        {
            NIR_Arrow.Visibility = IsActive(NIR_Header) ? Visibility.Collapsed : Visibility.Visible;
            NIR_Arrow1.Visibility = IsActive(Section1_Header) ? Visibility.Collapsed : Visibility.Visible;
            NIR_Arrow2.Visibility = IsActive(Section2_Header) ? Visibility.Collapsed : Visibility.Visible;
            NIR_Arrow3.Visibility = IsActive(Section3_Header) ? Visibility.Collapsed : Visibility.Visible;
            NIR_Arrow4.Visibility = IsActive(Section4_Header) ? Visibility.Collapsed : Visibility.Visible;
            NIR_Arrow5.Visibility = IsActive(Section5_Header) ? Visibility.Collapsed : Visibility.Visible;
            NIR_Arrow51.Visibility = IsActive(Item_51) ? Visibility.Collapsed : Visibility.Visible;
            NIR_Arrow53.Visibility = IsActive(Item_53) ? Visibility.Collapsed : Visibility.Visible;
            Study_Arrow.Visibility = IsActive(Study_Header) ? Visibility.Collapsed : Visibility.Visible;
            Study_Arrow14.Visibility = IsActive(Section14_Header) ? Visibility.Collapsed : Visibility.Visible;
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SearchPlaceholder.Visibility = string.IsNullOrEmpty(SearchBox.Text) ? Visibility.Visible : Visibility.Collapsed;
            string query = SearchBox.Text.ToLower();

            FilterMenuPanel(NIR_Menu, query);
            FilterMenuPanel(Study_Menu, query);
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
            TextBlock tb = null;
            if (border.Child is StackPanel sp)
                tb = sp.Children.OfType<TextBlock>().FirstOrDefault();
            else
                tb = border.Child as TextBlock;

            if (tb != null)
            {
                bool match = string.IsNullOrEmpty(query) || tb.Text.ToLower().Contains(query);
                border.Visibility = match ? Visibility.Visible : Visibility.Collapsed;
            }
        }
    }
}
