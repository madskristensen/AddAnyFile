namespace MadsKristensen.AddAnyFile.Templates
{
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    using M = TemplateMap.TemplateReference;

    /// <summary>
    /// Map of input patterns and their corresponding VS Templates. Ideally, these patterns would be configurable via the 
    /// Visual Studio Options dialog box.
    /// </summary>
    class TemplateMap : Dictionary<Regex, TemplateMap.TemplateReference>
    {
        /// <summary>
        /// Loads the default mappings.
        /// </summary>
        public void LoadDefaultMappings()
        {
            this[new Regex("c\\:(?<name>.*)") ] = new M { Language = "CSharp", TemplateName = "Class" };
            this[new Regex("i\\:(?<name>.*)") ] = new M { Language = "CSharp", TemplateName = "Interface" };
            this[new Regex("(?<name>.*)\\.js")] = new M { Language = "JavaScript", TemplateName = "JScript" };

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
            this[new Regex("(?<name>.*)\\.cs")] = new M { Language = "CSharp", TemplateName = "CodeFile" };
        }

        public TemplateMap()
        {
        }

        public class TemplateReference
        {
            public string Language { get; set; }
            public string TemplateName { get; set; }
        }
    }
}
