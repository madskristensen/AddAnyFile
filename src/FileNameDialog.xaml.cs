using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace MadsKristensen.AddAnyFile
{
    public partial class FileNameDialog : Window
    {
        private const string DEFAULT_TEXT = "Enter a file name";
        private static List<string> _tips = new List<string> {
            "Tip: 'folder/file.ext' also creates a new folder for the file",
            "Tip: You can create files starting with a dot, like '.gitignore'",
        };

        public FileNameDialog(string folder)
        {
            InitializeComponent();

            lblFolder.Content = string.Format("{0}/", folder);
            lblFolder.Loaded += delegate
            {
                txtName.Margin = new Thickness(10 + lblFolder.ActualWidth, 0, 100, 45);
                txtName.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
            };

            Loaded += (s, e) =>
            {
                Icon = BitmapFrame.Create(new Uri("pack://application:,,,/AddAnyFile;component/Resources/icon.png", UriKind.RelativeOrAbsolute));
                SetRandomTip();

                txtName.Focus();
                txtName.CaretIndex = 0;
                txtName.Text = DEFAULT_TEXT;

                txtName.PreviewKeyDown += delegate
                {
                    if (txtName.Text == DEFAULT_TEXT)
                        txtName.Text = string.Empty;
                };
            };

            PreviewKeyDown += (a, b) =>
            {
                if (b.Key == Key.Escape)
                    if (string.IsNullOrWhiteSpace(txtName.Text) || txtName.Text == DEFAULT_TEXT)
                        Close();
                    else
                        txtName.Text = string.Empty;
            };
        }

        private void SetRandomTip()
        {
            Random rnd = new Random(DateTime.Now.GetHashCode());
            int index = rnd.Next(_tips.Count);
            lblTips.Content = _tips[index];
        }

        public string Input
        {
            get { return txtName.Text.Trim(); }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}
