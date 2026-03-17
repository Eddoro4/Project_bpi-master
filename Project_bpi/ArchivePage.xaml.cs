using System.Windows;
using System.Windows.Controls;

namespace Project_bpi
{
    public partial class ArchivePage : UserControl
    {
        public ArchivePage()
        {
            InitializeComponent();
        }

        private void Filter_Click(object sender, RoutedEventArgs e)
        {
            FilterPopup.IsOpen = true;
        }

        private void ApplyFilter_Click(object sender, RoutedEventArgs e)
        {
            ApplyFilters();
            FilterPopup.IsOpen = false;
        }

        private void CancelFilter_Click(object sender, RoutedEventArgs e)
        {
            FilterPopup.IsOpen = false;
        }

        private void ApplyFilters()
        {
            string searchText = FilterTextBox.Text;
            string reportType = "Все";

            if (RadioResearch.IsChecked == true)
            {
                reportType = "Отчет по НИР";
            }
            else if (RadioEducational.IsChecked == true)
            {
                reportType = "Учебный отчет";
            }

            MessageBox.Show($"Применен фильтр:\nТип отчета: {reportType}\nПоиск: {searchText}",
                "Фильтр применен",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}
