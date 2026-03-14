using Project_bpi.Models;
using Project_bpi.Services;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace Project_bpi.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly DataBase _db = new DataBase();

        public ObservableCollection<Template> Templates { get; } = new ObservableCollection<Template>();

        private Template _selectedTemplate;
        public Template SelectedTemplate
        {
            get => _selectedTemplate;
            set { _selectedTemplate = value; OnPropertyChanged(); }
        }

        private object _currentView;
        public object CurrentView
        {
            get => _currentView;
            set { _currentView = value; OnPropertyChanged(); }
        }

        public ICommand ShowNIRReportsCommand { get; }
        public ICommand ShowTemplatesCommand { get; }

        public MainViewModel()
        {
            _db.InitializeDatabase();
            LoadTemplates();

            ShowNIRReportsCommand = new RelayCommand(_ => ShowNIRReports());
            ShowTemplatesCommand = new RelayCommand(_ => ShowTemplates());
        }

        private void LoadTemplates()
        {
            Templates.Clear();
            var list = _db.GetAllTemplatesWithHierarchy();
            foreach (var t in list)
                Templates.Add(t);
        }

        private void ShowTemplates()
        {
            // Создаём представление шаблонов и передаём в CurrentView
            var vm = new TemplatesViewModel(Templates);
            CurrentView = vm;
        }

        private void ShowNIRReports()
        {
            // Для демонстрации: отображаем список секций и подсекций первого шаблона, если он есть
            if (Templates.Any())
            {
                var template = Templates.First();
                var vm = new NIRReportsViewModel(template);
                CurrentView = vm;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
