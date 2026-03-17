using System.Windows;

namespace Project_bpi
{
    public partial class TemplateNameDialog : Window
    {
        public TemplateNameDialog()
        {
            InitializeComponent();
            Title = "\u041D\u043E\u0432\u043E\u0435 \u0438\u043C\u044F";
            Prompt = "\u0412\u0432\u0435\u0434\u0438\u0442\u0435 \u043D\u0430\u0437\u0432\u0430\u043D\u0438\u0435";
            Label = "\u041D\u0430\u0437\u0432\u0430\u043D\u0438\u0435";
            ConfirmButtonText = "\u0413\u043E\u0442\u043E\u0432\u043E";
            Loaded += (sender, args) => TemplateNameTextBox.Focus();
        }

        public string TemplateName
        {
            get => TemplateNameTextBox.Text;
            set => TemplateNameTextBox.Text = value;
        }

        public string Prompt
        {
            get => PromptTextBlock.Text;
            set => PromptTextBlock.Text = value;
        }

        public string Label
        {
            get => LabelTextBlock.Text;
            set => LabelTextBlock.Text = value;
        }

        public string ConfirmButtonText
        {
            get => ConfirmButton.Content?.ToString();
            set => ConfirmButton.Content = value;
        }

        private void Confirm_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TemplateName))
            {
                MessageBox.Show("\u0412\u0432\u0435\u0434\u0438\u0442\u0435 \u043d\u0430\u0437\u0432\u0430\u043d\u0438\u0435 \u0448\u0430\u0431\u043b\u043e\u043d\u0430.", "\u0428\u0430\u0431\u043b\u043e\u043d",
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
