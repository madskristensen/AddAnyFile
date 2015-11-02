using System.Text.RegularExpressions;
using EnvDTE80;

namespace MadsKristensen.AddAnyFile.Templates
{
    class ProjectItemManager
    {
        private readonly TemplateMap _templateMap;
        private readonly DTE2 _dte;

        public IItemCreator GetCreator(string rootFolder, string itemName)
        {
            string fileName;
            var template = GetTemplate(itemName, out fileName);

            if (template != null)
            {
                return new TemplatedItemCreator((Solution2)_dte.Solution, 
                    template, 
                    fileName, 
                    System.IO.Path.Combine(rootFolder + itemName));
            }
            else
            {
                return new NonTemplatedItemCreator(_dte, rootFolder, itemName);
            }
        }

        private TemplateMap.TemplateReference GetTemplate(string itemName, out string suggestedFileName)
        {
            Match match;
            foreach (var key in _templateMap.Keys)
                if ((match = key.Match(itemName)).Success && match.Groups.Count > 0)
                {
                    suggestedFileName = match.Groups["name"].Value;
                    return _templateMap[key];
                }

            suggestedFileName = null;
            return null;
        }

        public ProjectItemManager(DTE2 dte, TemplateMap templateMap)
        {
            _templateMap = templateMap;
            _dte = dte;
        }

        public class ItemBuilderToken
        {
            private readonly string _suggestedFileName;

            public ItemBuilderToken(string token)
            {
                _suggestedFileName = token;
            }
        }
    }
}
