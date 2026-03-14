using Project_bpi.Models;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Project_bpi.ViewModels
{
    public class NIRReportsViewModel : INotifyPropertyChanged
    {
        public Template Template { get; }
        public ObservableCollection<TemplateSection> Sections { get; } = new ObservableCollection<TemplateSection>();

        public NIRReportsViewModel(Template template)
        {
            Template = template;
            if (template.Sections != null)
            {
                foreach (var s in template.Sections)
                    Sections.Add(s);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
