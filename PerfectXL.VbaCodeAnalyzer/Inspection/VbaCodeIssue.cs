// Copyright 2017 Infotron B.V.
//
// This file is part of PerfectXL.VbaCodeAnalyzer.
// 
// PerfectXL.VbaCodeAnalyzer is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// PerfectXL.VbaCodeAnalyzer is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with PerfectXL.VbaCodeAnalyzer.  If not, see <http://www.gnu.org/licenses/>.

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
