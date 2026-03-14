using System;
using System.Windows;

namespace Project_bpi
{
    public partial class NewTemplateDialog : Window
    {
        public string TemplateName { get; private set; }
        public int Year { get; private set; }

        public NewTemplateDialog()
        {
            InitializeComponent();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(NameTextBox.Text))
            {
                MessageBox.Show("Введите имя шаблона", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            int year;
            if (!int.TryParse(YearTextBox.Text, out year) || year < 1900 || year > 3000)
            {
                MessageBox.Show("Введите корректный год", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            TemplateName = NameTextBox.Text.Trim();
            Year = year;
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
