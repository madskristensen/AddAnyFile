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

        public ItemInfo Create(Project project)
        {
            string file = Path.Combine(_folder, Path.Combine(_relativePath));
            string dir = Path.GetDirectoryName(file);

            try
            {
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Error creating the folder: " + dir);
                return ItemInfo.Empty;
            }

            if (!File.Exists(file))
            {
                int position = WriteFile(file);

                try
                {
                    if (AddFileToActiveProject(project, file))
                    {
                        Window window = _dte.ItemOperations.OpenFile(file);

                        // Move cursor into position
                        if (position > 0)
                        {
                            TextSelection selection = (TextSelection)window.Selection;
                            selection.CharRight(Count: position - 1);
                        }

                        return new ItemInfo
                        {
                            Extension = Path.GetExtension(file),
                            FileName = file
                        };
                    }

                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.Write(ex);
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The file '" + file + "' already exist.");
            }
            return ItemInfo.Empty;
        }

        private static bool AddFileToActiveProject(Project project, string fileName)
        {
            if (project == null || project.Kind == "{8BB2217D-0F2D-49D1-97BC-3654ED321F3B}") // ASP.NET 5 projects
                return false;

            string projectFilePath = AddAnyFilePackage.GetProjectRoot(project).Value.ToString();
            string projectDirPath = Path.GetDirectoryName(projectFilePath);

            if (!fileName.StartsWith(projectDirPath, StringComparison.OrdinalIgnoreCase))
                return false;

            var pi = project.ProjectItems.AddFromFile(fileName);
            if (fileName.EndsWith("__dummy__", StringComparison.Ordinal))
            {
                pi.Delete();
                return false;
            }
            else
            {
                return true;
            }
        }

        private static int WriteFile(string file)
        {
            Encoding encoding = new UTF8Encoding(true);
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
