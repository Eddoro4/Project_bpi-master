using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Project_bpi
{
  //архив
    public partial class ArchivePage : UserControl
    {
        public ArchivePage()
        {
            InitializeComponent();
        }

        private void Filter_Click(object sender, RoutedEventArgs e)
        {
            // Показываем popup при нажатии на кнопку фильтра
            FilterPopup.IsOpen = true;
        }

        private void ApplyFilter_Click(object sender, RoutedEventArgs e)
        {
            // Применяем фильтры
            ApplyFilters();

            // Закрываем popup
            FilterPopup.IsOpen = false;
        }

        private void CancelFilter_Click(object sender, RoutedEventArgs e)
        {
            // Закрываем popup без применения фильтров
            FilterPopup.IsOpen = false;
        }

        private void ApplyFilters()
        {
            // Здесь реализуйте логику применения фильтров
            string searchText = FilterTextBox.Text;
            string reportType = "Все";

            if (RadioResearch.IsChecked == true)
                reportType = "Отчет по НИР";
            else if (RadioEducational.IsChecked == true)
                reportType = "Учебный отчет";

            // Пример применения фильтра - в реальном приложении здесь будет фильтрация данных
            MessageBox.Show($"Применен фильтр:\nТип отчета: {reportType}\nПоиск: {searchText}",
                          "Фильтр применен",
                          MessageBoxButton.OK,
                          MessageBoxImage.Information);
        }
    }
}