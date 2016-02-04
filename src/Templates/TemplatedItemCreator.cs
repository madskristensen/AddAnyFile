namespace MadsKristensen.AddAnyFile.Templates
{
    using System;
    using System.Linq;
    using System.Xml.Linq;
    using EnvDTE;
    using EnvDTE80;

    class TemplatedItemCreator : IItemCreator
    {
        private readonly TemplateMapping _template;
        private readonly Solution2 _solution;
        private readonly string[] _relativePath;

        public TemplatedItemCreator(Solution2 solution,
                TemplateMapping template,
                string[] relativePath)
        {
            this._solution = solution;
            this._template = template;
            this._relativePath = relativePath;
        }

        public ItemInfo Create(Project project)
        {
            string templatePath = _solution.GetProjectItemTemplate(_template.TemplateName, _template.Language);
            string ext = GetTargetExtension(templatePath);

            string itemName = _relativePath[_relativePath.Length - 1];
            if (!itemName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                itemName += ext;

            var parent = GetItemParent(project);
            parent.AddFromTemplate(templatePath, itemName);

            return new ItemInfo {
                Extension = ext,
                FileName = itemName
            };
        }

        private ProjectItems GetItemParent(Project project)
        {
            var parent = project.ProjectItems;
            Func<string, ProjectItems, ProjectItem> existingItem = (name, items) => items
                                                        .Cast<ProjectItem>()
                                                        .FirstOrDefault(p => p.Name == name);

            for (int i = 0; i < _relativePath.Length - 1; i++)
            {
                var p = existingItem(_relativePath[i], parent);
                if (p == null)
                    parent = parent.AddFolder(_relativePath[i]).ProjectItems;
                else
                    parent = p.ProjectItems;
            }
            return parent;
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
