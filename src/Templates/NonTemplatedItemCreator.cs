using System;
using System.IO;
using System.Reflection;
using System.Text;
using EnvDTE;
using EnvDTE80;

namespace MadsKristensen.AddAnyFile.Templates
{
    class NonTemplatedItemCreator : IItemCreator
    {
        private readonly string[] _relativePath;
        private readonly string _folder;
        private readonly DTE2 _dte;

        public NonTemplatedItemCreator(DTE2 dte, string rootFolder, string[] relativePath)
        {
            this._relativePath = relativePath;
            this._folder = rootFolder;
            this._dte = dte;
        }

        public void Create(Project project)
        {
            string file = Path.Combine(_folder, _relativePath[_relativePath.Length - 1]);
            string dir = Path.GetDirectoryName(file);

            try
            {
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Error creating the folder: " + dir);
                return;
            }

            if (!File.Exists(file))
            {
                int position = WriteFile(file);

                try
                {
                    AddFileToActiveProject(project, file);
                    Window window = _dte.ItemOperations.OpenFile(file);

                    // Move cursor into position
                    if (position > 0)
                    {
                        TextSelection selection = (TextSelection)window.Selection;
                        selection.CharRight(Count: position - 1);
                    }

                }
                catch { /* Something went wrong. What should we do about it? */ }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The file '" + file + "' already exist.");
            }
        }

        private static void AddFileToActiveProject(Project project, string fileName)
        {
            if (project == null || project.Kind == "{8BB2217D-0F2D-49D1-97BC-3654ED321F3B}") // ASP.NET 5 projects
                return;

            string projectFilePath = AddAnyFilePackage.GetProjectRoot(project).Value.ToString();
            string projectDirPath = Path.GetDirectoryName(projectFilePath);

            if (!fileName.StartsWith(projectDirPath, StringComparison.OrdinalIgnoreCase))
                return;

            project.ProjectItems.AddFromFile(fileName);
        }

        private static int WriteFile(string file)
        {
            Encoding encoding = new UTF8Encoding(false);
            string extension = Path.GetExtension(file);

            string assembly = Assembly.GetExecutingAssembly().Location;
            string folder = Path.GetDirectoryName(assembly).ToLowerInvariant();
            string template = Path.Combine(folder, "Templates\\", extension);

            if (File.Exists(template))
            {
                string content = File.ReadAllText(template);
                int index = content.IndexOf('$');
                content = content.Remove(index, 1);
                File.WriteAllText(file, content, encoding);
                return index;
            }

            File.WriteAllText(file, string.Empty, encoding);
            return 0;
        }
    }
}
