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
            // Создаём файл в подпапке "templates" приложения
            var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "templates");
            Directory.CreateDirectory(dir);
            var filename = GenerateUniqueFileName(dir, "template", ".docx");
            var fullpath = Path.Combine(dir, filename);

            // Создать пустой Word документ
            _db.CreateEmptyWordDocument(fullpath);

            // Создать запись в БД
            var template = new Template
            {
                Name = Path.GetFileNameWithoutExtension(filename),
                Year = DateTime.Now.Year,
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
