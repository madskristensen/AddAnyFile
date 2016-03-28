using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio.Shell;

namespace MadsKristensen.AddAnyFile
{
    static class TemplateMap
    {
        static string _folder;
        static readonly string[] _templateFiles;
        const string _defaultExt = ".txt";

        static TemplateMap()
        {
            string assembly = Assembly.GetExecutingAssembly().Location;
            _folder = Path.GetDirectoryName(assembly).ToLowerInvariant() + "\\Templates";
            _templateFiles = Directory.GetFiles(_folder, "*" + _defaultExt, SearchOption.AllDirectories);
        }

        public static async Task<string> GetTemplateFilePath(Project project, string file)
        {
            string extension = Path.GetExtension(file).ToLowerInvariant();
            string name = Path.GetFileName(file);
            string safeName = Path.GetFileNameWithoutExtension(file);
            string relative = PackageUtilities.MakeRelative(project.GetRootFolder(), Path.GetDirectoryName(file));

            string templateFile = null;

            // Look for direct file name matches
            if (_templateFiles.Any(f => Path.GetFileName(f).Equals(name + _defaultExt, StringComparison.OrdinalIgnoreCase)))
            {
                templateFile = GetTemplate(name);
            }

            // Look for file extension matches
            else if (_templateFiles.Any(f => Path.GetFileName(f).Equals(extension + _defaultExt, StringComparison.OrdinalIgnoreCase)))
            {
                var tmpl = AdjustForSpecific(safeName, extension);
                templateFile = GetTemplate(tmpl);
            }

            return await ReplaceTokens(project, safeName, relative, templateFile);
        }

        private static string GetTemplate(string name)
        {
            return Path.Combine(_folder, name + _defaultExt);
        }

        private static async Task<string> ReplaceTokens(Project project, string name, string relative, string templateFile)
        {
            if (string.IsNullOrEmpty(templateFile))
                return templateFile;

            string rootNs = project.GetRootNamespace();
            string ns = string.IsNullOrEmpty(rootNs) ? "MyNamespace" : rootNs;

            if (!string.IsNullOrEmpty(relative))
            {
                ns += "." + ProjectHelpers.CleanNameSpace(relative);
            }

            using (var reader = new StreamReader(templateFile))
            {
                string content = await reader.ReadToEndAsync();

                return content.Replace("{namespace}", ns)
                              .Replace("{itemname}", name);
            }
        }

        private static string AdjustForSpecific(string safeName, string extension)
        {
            if (Regex.IsMatch(safeName, "^I[A-Z].*"))
                return extension += "-interface";

            return extension;
        }
    }
}
