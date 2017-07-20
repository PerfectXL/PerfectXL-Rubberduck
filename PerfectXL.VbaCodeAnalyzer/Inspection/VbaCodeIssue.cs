using System;
using System.Linq;
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
            var result = "";
            var count = text.Count(f => f == (char)39);

            //result = text.Contains("Option Explicit")
            //    ? "Option Explicit"
            //    : text.Substring(text.IndexOf((char)39), text.LastIndexOf((char)39) - text.IndexOf((char)39));

            switch (count)
            {
                case 0:
                case 1:
                    result = text.Contains("Option Explicit")? "Option Explicit": "";
                    break;
                case 2:
                    result = text.Substring(text.IndexOf((char)39), text.LastIndexOf((char)39) + 1- text.IndexOf((char)39));
                    break;
            }
            return result;
        }
    }
}