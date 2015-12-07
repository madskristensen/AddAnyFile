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
            "Tip: You can create files without file extensions, like 'LICENSE'",
        };

        public FileNameDialog(string folder, string defaultExt)
        {
            InitializeComponent();

            lblFolder.Content = string.Format("{0}/", folder);

            Loaded += (s, e) =>
            {
                Icon = BitmapFrame.Create(new Uri("pack://application:,,,/AddAnyFile;component/Resources/icon.png", UriKind.RelativeOrAbsolute));
                SetRandomTip();

                txtName.Focus();
                txtName.CaretIndex = 0;
                txtName.Text = "filename" + defaultExt;
                txtName.Select(0, txtName.Text.Length - defaultExt.Length);
                /*DEFAULT_TEXT;

                txtName.PreviewKeyDown += delegate
                {
                    if (txtName.Text == DEFAULT_TEXT)
                    {
                        txtName.Text = string.Empty;
                        btnCreate.IsEnabled = true;
                    }
                };
                */
            };

            PreviewKeyDown += (a, b) =>
            {
                int p = -1;

                if (b.Key == Key.Escape)
                {
                    if (string.IsNullOrWhiteSpace(txtName.Text) || txtName.Text == DEFAULT_TEXT)
                        Close();
                    else
                        txtName.Text = string.Empty;
                }
                else if (b.Key == Key.OemPeriod && (p = txtName.Text.LastIndexOf('.')) > -1)
                {
                    txtName.Select(p, txtName.Text.Length - p);
                }
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
