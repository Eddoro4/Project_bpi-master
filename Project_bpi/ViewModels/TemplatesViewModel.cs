using Project_bpi.Models;
using System.Collections.ObjectModel;

namespace Project_bpi.ViewModels
{
    public class TemplatesViewModel
    {
        public ObservableCollection<Template> Templates { get; }

        public TemplatesViewModel(ObservableCollection<Template> templates)
        {
            Templates = templates;
        }
    }
}
