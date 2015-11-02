namespace MadsKristensen.AddAnyFile.Templates
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Xml.Linq;
    using EnvDTE;
    using EnvDTE80;

    class TemplatedItemCreator : IItemCreator
    {
        private readonly string _fileName;
        private readonly TemplateMap.TemplateReference _template;
        private readonly Solution2 _solution;
        private readonly string _fullName;

        public TemplatedItemCreator(Solution2 solution, 
                TemplateMap.TemplateReference template, 
                string fileName,
                string givenName)
        {
            this._solution = solution;
            this._template = template;
            this._fileName = fileName;
            this._fullName = givenName;
        }

        public void Create(Project project)
        {
            string templatePath = _solution.GetProjectItemTemplate(_template.TemplateName, _template.Language);
            string ext = GetTargetExtension(templatePath);

            string itemName = System.IO.Path.GetFileName(_fileName);
            if (!_fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                itemName += ext;

            var parent = GetItemParent(project);
            parent.AddFromTemplate(templatePath, itemName);
        }

        private ProjectItems GetItemParent(Project project)
        {
            string[] directories = GetRelativePath(project, _fullName).Split('\\');

            var parent = project.ProjectItems;
            Func<string, ProjectItems, ProjectItem> existingItem = (name, items) => items
                                                        .Cast<ProjectItem>()
                                                        .FirstOrDefault(p => p.Name == name);

            for (int i = 0; i < directories.Length - 1; i++)
            {
                var p = existingItem(directories[i], parent);
                if (p == null)
                    parent = parent.AddFolder(directories[i]).ProjectItems;
                else
                    parent = p.ProjectItems;
            }
            return parent;
        }

        private static string GetRelativePath(Project project, string fullName)
        {
            string projectPath = Path.GetDirectoryName(project.FullName);
            if (fullName.StartsWith(projectPath, StringComparison.OrdinalIgnoreCase))
                return fullName.Substring(projectPath.Length + 1);
            else
                return fullName;
        }

        private string GetTargetExtension(string templateFile)
        {
            var doc = XDocument.Load(templateFile);
            XNamespace ns = "http://schemas.microsoft.com/developer/vstemplate/2005";
            string defaultName = doc.Document.Descendants(ns + "TemplateData")
                                    .First()
                                    .Descendants(ns + "DefaultName")
                                    .First()
                                    .Value;
            return System.IO.Path.GetExtension(defaultName);

        }
    }
}
