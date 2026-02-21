using System.Windows.Controls;

namespace Project_bpi
{
    public partial class Item53_View : UserControl
    {
        public Item53_View()
        {
            InitializeComponent();
        }

        private void EditButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ContentTextBox.IsReadOnly = false;
            ContentTextBox.Focus();
            SaveButton.IsEnabled = true; // Активируем кнопку сохранения
        }

        private void SaveButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            // 🔽 Здесь — ваша логика сохранения
            // Например: DataManager.SaveSection("Cooperation", ContentTextBox.Text);

            ContentTextBox.IsReadOnly = true;
            SaveButton.IsEnabled = false; // Деактивируем после сохранения

            // Опционально: можно уведомить пользователя об успехе
        }
    }
}