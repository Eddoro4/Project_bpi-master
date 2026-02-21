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
    public partial class TemplatesPage : UserControl
    {
        public TemplatesPage()
        {
            InitializeComponent();
        }

        private void TemplateButton_Click(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            if (button != null)
            {
                string templateName = button.Content.ToString();

                // Здесь можно добавить логику загрузки/применения шаблона
                MessageBox.Show($"Выбран шаблон: {templateName}\n\nШаблон будет загружен и применен.",
                              "Шаблон выбран",
                              MessageBoxButton.OK,
                              MessageBoxImage.Information);

                // Пример вызова метода загрузки шаблона:
                // LoadTemplate(templateName);
            }
        }
        private void EditTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            // Логика редактирования шаблона
            Button button = sender as Button;
            if (button != null)
            {
                // Можно добавить логику для определения, какой шаблон редактируется
                MessageBox.Show("Редактирование шаблона");
            }
        }

        private void CreateTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            // Логика создания нового шаблона
            MessageBox.Show("Создание нового шаблона");
        }
    }
}