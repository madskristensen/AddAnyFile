﻿using System.IO;
using System.Text.RegularExpressions;
using EnvDTE80;

namespace MadsKristensen.AddAnyFile.Templates
{
    class ProjectItemManager
    {
        private readonly TemplateMap _templateMap;
        private readonly DTE2 _dte;

        public IItemCreator GetCreator(string rootFolder, string relativePath)
        {
            var path = relativePath.Split(Path.DirectorySeparatorChar);
            string fileName;
            var template = GetTemplate(path[path.Length-1], out fileName);

            if (template != null)
            {
                path[path.Length - 1] = fileName;
                return new TemplatedItemCreator((Solution2)_dte.Solution, 
                    template, 
                    rootFolder,
                    path);
            }
            else
            {
                return new NonTemplatedItemCreator(_dte, rootFolder, relativePath);
            }
        }

        private TemplateMapping GetTemplate(string itemName, out string suggestedFileName)
        {
            Match match;
            foreach (var mapping in _templateMap)
                if ((match = mapping.GetPatternExpression().Match(itemName)).Success && match.Groups.Count > 0)
                {
                    suggestedFileName = match.Groups["name"].Value;
                    return mapping;
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
