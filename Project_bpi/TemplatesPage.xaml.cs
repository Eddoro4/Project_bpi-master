using Project_bpi.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
            this.DataContext = this;
            UpdateTemplates();
        }
        private ObservableCollection<Template> _templates;
        public ObservableCollection<Template> Templates { 
            get { return _templates; }
            set
            {
                _templates = value;
                
            }
        }
        private void UpdateTemplates()
        {
            var temp = App.DB.GetTemplates();
            Templates = new ObservableCollection<Template>(temp);
        }
    }
}