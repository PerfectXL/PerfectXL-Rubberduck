using System;
using Rubberduck.Parsing.Inspections.Abstract;

namespace infotron.VbaCodeAnalizer.Inspections {
    public class VbaCodeIssue
    {
        public VbaCodeIssue(IInspectionResult item)
        {
            Severity = item.Inspection.Severity.ToString();
            Description = item.Description;
            Type = item.Inspection.AnnotationName;
            Meta = item.Inspection.Meta;
            Name = ExtractIdentifierName(item.Description);
            Line = item.QualifiedSelection.Selection.StartLine;
            Column = item.QualifiedSelection.Selection.StartColumn;
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
            return text.Contains("Option Explicit")
                ? "Option Explicit"
                : text.Substring(text.IndexOf("'", StringComparison.Ordinal),
                    text.LastIndexOf("'", StringComparison.Ordinal) - text.IndexOf("'", StringComparison.Ordinal));
        }
    }
}