using System.Windows;

namespace Project_bpi
{
    public partial class ArchivePage
    {
        private void ApplyFilterFixed_Click(object sender, RoutedEventArgs e)
        {
            string searchText = FilterTextBox.Text;
            string reportType = "\u0412\u0441\u0435";

            if (RadioResearch.IsChecked == true)
            {
                reportType = "\u041e\u0442\u0447\u0435\u0442 \u043f\u043e \u041d\u0418\u0420";
            }
            else if (RadioEducational.IsChecked == true)
            {
                reportType = "\u0423\u0447\u0435\u0431\u043d\u044b\u0439 \u043e\u0442\u0447\u0435\u0442";
            }

            MessageBox.Show(
                $"\u041f\u0440\u0438\u043c\u0435\u043d\u0435\u043d \u0444\u0438\u043b\u044c\u0442\u0440:\n\u0422\u0438\u043f \u043e\u0442\u0447\u0435\u0442\u0430: {reportType}\n\u041f\u043e\u0438\u0441\u043a: {searchText}",
                "\u0424\u0438\u043b\u044c\u0442\u0440 \u043f\u0440\u0438\u043c\u0435\u043d\u0435\u043d",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            FilterPopup.IsOpen = false;
        }
    }
}
