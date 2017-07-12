using System.Text.RegularExpressions;
using Rubberduck.Parsing.Inspections.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Inspection
{
    public class VbaCodeIssue
    {
        public VbaCodeIssue(IInspectionResult item, string fileName, string moduleName)
        {
            Severity = item.Inspection.Severity.ToString();
            Description = item.Description;
            Type = item.Inspection.AnnotationName;
            Meta = item.Inspection.Meta;
            Name = ExtractIdentifierName(item.Description);
            Line = item.QualifiedSelection.Selection.StartLine;
            Column = item.QualifiedSelection.Selection.StartColumn;
            FileName = fileName;
            ModuleName = moduleName;
        }

        public string Type { get; set; }
        public string ModuleName { get; set; }
        public string Severity { get; set; }
        public string Description { get; set; }
        public string Name { get; set; }
        public string Meta { get; set; }
        public int Line { get; set; }
        public int Column { get; set; }
        public string FileName { get; set; }

        private static string ExtractIdentifierName(string text)
        {
            if (text.Contains("Option Explicit"))
            {
                return "Option Explicit";
            }
            Match match = Regex.Match(text, @" ['‘’] ( [^'‘’]+ ) ['‘’] ", RegexOptions.IgnorePatternWhitespace);
            return match.Success ? match.Groups[1].Value : text;
        }
    }
}