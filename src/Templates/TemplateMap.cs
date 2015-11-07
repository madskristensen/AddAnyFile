namespace MadsKristensen.AddAnyFile.Templates
{
    using System.Collections.Generic;
    using System.IO;
    using System.Text.RegularExpressions;
    using Newtonsoft.Json;
    using M = TemplateMap.TemplateMapping;

    /// <summary>
    /// Map of input patterns and their corresponding VS Templates. Ideally, these patterns would be configurable via the 
    /// Visual Studio Options dialog box.
    /// </summary>
    class TemplateMap : List<TemplateMap.TemplateMapping>
    {
        /// <summary>
        /// Loads the default mappings.
        /// </summary>
        public void LoadDefaultMappings()
        {
            this.Add(new M(@"^c\:(?<name>.*)", "CSharp", "Class"));
            this.Add(new M(@"^i\:(?<name>.*)", "CSharp", "Interface"));
            this.Add(new M(@"^(?<name>I.*)\.cs$", "CSharp", "Interface"));
            this.Add(new M(@"^(?<name>I.*)\.vb$", "VisualBasic", "Interface"));
            /* NOTE: With VS 2015 Community Edition, CodeFile.cs was not replacing the $rootnamespace$ tag.
             * If this happens, open the CodeFile.vstemplate file 
             * (e.g. in C:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\ItemTemplates\CSharp\Code\1033\CodeFile\CodeFile.vstemplate)
             * and replace
             
              <TemplateContent>
                <ProjectItem>CodeFile.cs</ProjectItem>
              </TemplateContent>

            * with

              <TemplateContent>
                <ProjectItem ReplaceParameters="true">CodeFile.cs</ProjectItem>
              </TemplateContent>

            */
            this.Add(new M(@"^e:(?<name>.*)\.cs", "CSharp", "CodeFile"));
            // Following mappings are the broadest from a language POV. So, they need to be at the end of the mappings
            this.Add(new M(@"^(?<name>.*)\.cs$", "CSharp", "Class"));
            this.Add(new M(@"^(?<name>.*)\.js$", "JavaScript", "JScript"));
        }

        public static TemplateMap LoadFromFile(string path)
        {
            if (File.Exists(path))
            {
                try
                {
                    string contents = File.ReadAllText(path);
                    return JsonConvert.DeserializeObject<TemplateMap>(contents);
                }
                catch (Newtonsoft.Json.JsonSerializationException e)
                {
                    AddAnyFilePackage.LogToOutputPane(string.Concat(
                        "There was error loading the mappings from: ",
                        path,
                        "\r\n",
                        e.Message));
                }
            }
            return null;
        }

        public static void WriteToFile(TemplateMap map, string path)
        {
            try
            {
                var json = JsonConvert.SerializeObject(map, Formatting.Indented);
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, json, System.Text.Encoding.UTF8);
            }
            catch (System.Exception e)
            {
                AddAnyFilePackage.LogToOutputPane(string.Concat(
                    "Could not save the mapping file to: ",
                    path,
                    "\r\n",
                    e.Message));

            }
        }

        public TemplateMap()
        {
        }

        public class TemplateMapping
        {
            public string Pattern { get; set; }
            public string Language { get; set; }
            public string TemplateName { get; set; }

            private Regex _patternExpression;
            private readonly object _lock = new object();

            public Regex GetPatternExpression()
            {
                if (_patternExpression == null)
                    lock (_lock)
                        if (_patternExpression == null)
                        {
                            _patternExpression = CreateExpression(Pattern);

                        }
                return _patternExpression;
            }

            private static Regex CreateExpression(string expression)
            {
                return new Regex(expression, RegexOptions.Compiled);
            }

            public TemplateMapping()
            {
            }

            public TemplateMapping(string pattern, string language, string templateName)
            {
                this.Pattern = pattern;
                this.Language = language;
                this.TemplateName = templateName;
                this._patternExpression = CreateExpression(pattern);                
            }

            public override string ToString()
            {
                return string.Concat(Pattern, " => ", Language, "/", TemplateName);
            }
        }
    }
}
