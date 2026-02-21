using System;
using System.Windows;
using System.Windows.Controls;

namespace Project_bpi
{
    public partial class HistoryChangesView : UserControl
    {
        public HistoryChangesView()
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
            bool includeEdit = CheckEdit.IsChecked == true;
            bool includeDelete = CheckDelete.IsChecked == true;
            string dateFilter = FilterDateBox.Text.Trim();

            if (!includeEdit && !includeDelete)
            {
                MessageBox.Show("Выберите хотя бы один тип действия: Редактирование или Удаление.",
                                "Фильтр",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                return;
            }

            string actions = "";
            if (includeEdit) actions += "Редактирование, ";
            if (includeDelete) actions += "Удаление, ";

            // Безопасное удаление последней запятой и пробела
            if (actions.EndsWith(", "))
                actions = actions.Substring(0, actions.Length - 2);

            string message = $"Применены фильтры:\nТипы: {actions}\nДата: {(string.IsNullOrEmpty(dateFilter) ? "любая" : dateFilter)}";
            MessageBox.Show(message, "Фильтр применён", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}