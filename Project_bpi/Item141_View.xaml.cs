using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Project_bpi
{
    public partial class Item141_View : UserControl
    {
        private bool _sortButtonsVisible = false;
        private int _currentColumnIndex = -1;
        private Button _currentSortButton;

        // === НОВАЯ ФУНКЦИОНАЛЬНОСТЬ ===
        private bool _isDeletionMode = false;
        private Brush _defaultDeleteButtonBackground;

        public Item141_View()
        {
            InitializeComponent();

            // === СТАРАЯ ФУНКЦИОНАЛЬНОСТЬ ===
            FilterButton.Click += OnFilterButtonClick;
            SetupSortButtonEvents();

            SortAscButton.Click += OnSortAscButtonClick;
            SortDescButton.Click += OnSortDescButtonClick;
            CancelButton.Click += OnCancelButtonClick;
            OkButton.Click += OnOkButtonClick;
            SearchTextBox.GotFocus += OnSearchTextBoxGotFocus;
            SearchTextBox.LostFocus += OnSearchTextBoxLostFocus;
            this.MouseDown += OnUserControlMouseDown;

            this.Loaded += (s, e) =>
            {
                FilterSortPanel.Visibility = Visibility.Collapsed;
            };

        }

        // === СТАРАЯ ЛОГИКА: ФИЛЬТРАЦИЯ И СОРТИРОВКА ===
        private void OnFilterButtonClick(object sender, RoutedEventArgs e)
        {
            _sortButtonsVisible = !_sortButtonsVisible;
            var visibility = _sortButtonsVisible ? Visibility.Visible : Visibility.Collapsed;

            SortCol0.Visibility = visibility;
            SortCol1.Visibility = visibility;
            SortCol2.Visibility = visibility;
            SortCol3.Visibility = visibility;
            SortCol4.Visibility = visibility;
            SortCol6.Visibility = visibility;
        }

        private void SetupSortButtonEvents()
        {
            SortCol0.Click += (s, e) => ShowFilterSortPanel(0, SortCol0);
            SortCol1.Click += (s, e) => ShowFilterSortPanel(1, SortCol1);
            SortCol2.Click += (s, e) => ShowFilterSortPanel(2, SortCol2);
            SortCol3.Click += (s, e) => ShowFilterSortPanel(3, SortCol3);
            SortCol4.Click += (s, e) => ShowFilterSortPanel(4, SortCol4);
            SortCol6.Click += (s, e) => ShowFilterSortPanel(6, SortCol6);
        }

        private void ShowFilterSortPanel(int columnIndex, Button sortButton)
        {
            _currentColumnIndex = columnIndex;
            _currentSortButton = sortButton;

            Point buttonPosition = sortButton.TransformToAncestor(this).Transform(new Point(0, 0));
            double panelWidth = 246;
            double panelHeight = 348;
            double left = buttonPosition.X + sortButton.ActualWidth + 5;
            double top = buttonPosition.Y - 20;

            if (left + panelWidth > this.ActualWidth)
                left = buttonPosition.X - panelWidth - 5;
            if (top + panelHeight > this.ActualHeight)
                top = this.ActualHeight - panelHeight - 10;
            if (top < 0)
                top = 10;

            FilterSortPanel.Margin = new Thickness(left, top, 0, 0);

            SortAscButton.Content = "Сортировать по";
            SortDescButton.Content = "Сортировать по";
            SearchResultsPanel.Children.Clear();
            SearchTextBox.Text = "Поиск...";
            FilterSortPanel.Visibility = Visibility.Visible;
        }

        private void SortColumn(int columnIndex, bool ascending)
        {
            // Здесь можно реализовать сортировку _contracts
            FilterSortPanel.Visibility = Visibility.Collapsed;
        }

        private void OnSortAscButtonClick(object sender, RoutedEventArgs e)
        {
            SortColumn(_currentColumnIndex, true);
        }

        private void OnSortDescButtonClick(object sender, RoutedEventArgs e)
        {
            SortColumn(_currentColumnIndex, false);
        }

        private void OnCancelButtonClick(object sender, RoutedEventArgs e)
        {
            FilterSortPanel.Visibility = Visibility.Collapsed;
        }

        private void OnOkButtonClick(object sender, RoutedEventArgs e)
        {
            FilterSortPanel.Visibility = Visibility.Collapsed;
        }

        private void OnSearchTextBoxGotFocus(object sender, RoutedEventArgs e)
        {
            if (SearchTextBox.Text == "Поиск...")
            {
                SearchTextBox.Text = "";
                SearchTextBox.Foreground = Brushes.Black;
            }
        }

        private void OnSearchTextBoxLostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SearchTextBox.Text))
            {
                SearchTextBox.Text = "Поиск...";
                SearchTextBox.Foreground = Brushes.Gray;
            }
        }



        private void OnUserControlMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (FilterSortPanel.Visibility == Visibility.Visible)
            {
                var originalSource = e.OriginalSource as DependencyObject;
                var isPartOfFilterPanel = false;
                while (originalSource != null)
                {
                    if (originalSource == FilterSortPanel)
                    {
                        isPartOfFilterPanel = true;
                        break;
                    }
                    originalSource = VisualTreeHelper.GetParent(originalSource);
                }
                if (!isPartOfFilterPanel)
                {
                    FilterSortPanel.Visibility = Visibility.Collapsed;
                }
            }
        }
    }
}