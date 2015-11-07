using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace MadsKristensen.AddAnyFile.Templates
{
    [DataContract]
    public class TemplateMapping
    {
        [DataMember]
        public string Pattern { get; set; }
        [DataMember]
        public string Language { get; set; }
        [DataMember]
        public string TemplateName { get; set; }

        private Regex _patternExpression;

        public Regex GetPatternExpression()
        {
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
