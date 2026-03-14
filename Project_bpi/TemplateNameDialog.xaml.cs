using System.Windows;

namespace Project_bpi
{
    public partial class TemplateNameDialog : Window
    {
        public TemplateNameDialog()
        {
            InitializeComponent();
            Loaded += (sender, args) => TemplateNameTextBox.Focus();
        }

        public string TemplateName
        {
            get => TemplateNameTextBox.Text;
            set => TemplateNameTextBox.Text = value;
        }

        private void Confirm_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TemplateName))
            {
                MessageBox.Show("Введите название шаблона.", "Шаблон",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
