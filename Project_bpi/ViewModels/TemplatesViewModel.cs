using Project_bpi.Models;
using Project_bpi.Services;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows.Input;
using System;

namespace Project_bpi.ViewModels
{
    public class TemplatesViewModel
    {
        private readonly DataBase _db = new DataBase();
        public ObservableCollection<Template> Templates { get; }
        public ICommand AddTemplateCommand { get; }

        public TemplatesViewModel(ObservableCollection<Template> templates)
        {
            Templates = templates;
            AddTemplateCommand = new RelayCommand(_ => AddTemplate());
        }

        private void AddTemplate()
        {
            // Создаём простой WPF-диалог на лету для ввода имени и года
            var owner = System.Windows.Application.Current?.MainWindow;
            var dlg = new System.Windows.Window
            {
                Title = "Новый шаблон",
                SizeToContent = System.Windows.SizeToContent.WidthAndHeight,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner,
                ResizeMode = System.Windows.ResizeMode.NoResize,
                Owner = owner
            };

            var grid = new System.Windows.Controls.Grid { Margin = new System.Windows.Thickness(10) };
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });

            var namePanel = new System.Windows.Controls.StackPanel { Margin = new System.Windows.Thickness(0, 0, 0, 8) };
            System.Windows.Controls.Grid.SetRow(namePanel, 0);
            namePanel.Children.Add(new System.Windows.Controls.TextBlock { Text = "Имя шаблона", FontWeight = System.Windows.FontWeights.SemiBold });
            var nameBox = new System.Windows.Controls.TextBox { Width = 280 };
            namePanel.Children.Add(nameBox);
            grid.Children.Add(namePanel);

            var yearPanel = new System.Windows.Controls.StackPanel { Margin = new System.Windows.Thickness(0, 0, 0, 8) };
            System.Windows.Controls.Grid.SetRow(yearPanel, 1);
            yearPanel.Children.Add(new System.Windows.Controls.TextBlock { Text = "Год", FontWeight = System.Windows.FontWeights.SemiBold });
            var yearBox = new System.Windows.Controls.TextBox { Width = 120, Text = DateTime.Now.Year.ToString() };
            yearPanel.Children.Add(yearBox);
            grid.Children.Add(yearPanel);

            var buttons = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Right };
            System.Windows.Controls.Grid.SetRow(buttons, 2);
            var cancelBtn = new System.Windows.Controls.Button { Content = "Отмена", Width = 80, Margin = new System.Windows.Thickness(0, 0, 8, 0) };
            var okBtn = new System.Windows.Controls.Button { Content = "ОК", Width = 80 };
            buttons.Children.Add(cancelBtn);
            buttons.Children.Add(okBtn);
            grid.Children.Add(buttons);

            cancelBtn.Click += (s, e) => { dlg.DialogResult = false; dlg.Close(); };
            okBtn.Click += (s, e) =>
            {
                var enteredName = nameBox.Text?.Trim();
                if (string.IsNullOrWhiteSpace(enteredName))
                {
                    System.Windows.MessageBox.Show("Введите имя шаблона", "Внимание", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }
                if (!int.TryParse(yearBox.Text, out int enteredYear) || enteredYear < 1900 || enteredYear > 3000)
                {
                    System.Windows.MessageBox.Show("Введите корректный год", "Внимание", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }
                // Сохраним данные в Tag и закроем диалог
                dlg.Tag = new Tuple<string, int>(enteredName, enteredYear);
                dlg.DialogResult = true;
                dlg.Close();
            };

            dlg.Content = grid;
            var res = dlg.ShowDialog();
            if (res != true) return;
            var tuple = dlg.Tag as Tuple<string, int>;
            if (tuple == null) return;
            var name = tuple.Item1;
            var year = tuple.Item2;

            // Создаём файл в подпапке "templates" приложения
            var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "templates");
            Directory.CreateDirectory(dir);
            var filename = GenerateUniqueFileName(dir, name, ".docx");
            var fullpath = Path.Combine(dir, filename);

            // Создать пустой Word документ
            _db.CreateEmptyWordDocument(fullpath);

            // Создать запись в БД
            var template = new Template
            {
                Name = name,
                Year = year,
                Path = fullpath
            };
            var id = _db.InsertTemplate(template);
            template.Id = id;

            // Обновить коллекцию в UI
            Templates.Add(template);
        }

        private string GenerateUniqueFileName(string dir, string baseName, string ext)
        {
            int i = 1;
            string name;
            do
            {
                name = $"{baseName}_{DateTime.Now:yyyyMMddHHmmss}_{i}{ext}";
                i++;
            } while (File.Exists(Path.Combine(dir, name)));
            return name;
        }
    }
}
