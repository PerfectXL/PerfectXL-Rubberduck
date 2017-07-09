using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using infotron.VbaCodeAnalizer.Mog;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;


namespace infotron.VbaCodeAnalizer.Inspections
{
    public static class Inspector
    {
        public static List<Issue> Inspect<TInspection>(string inputcode, Helper helper) where TInspection : InspectionBase
        {
            IEnumerable<IInspectionResult> inspectionResults = null;

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out component);
            var state = MockParser.CreateAndParse(vbe);

            var inspection = InspectionFactory.Create<TInspection>(state);

            switch (helper)
            {
                case Helper.Yes:
                    var inspector = InspectionsHelper.GetInspector(inspection);
                    inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                    break;
                case Helper.No:
                    inspectionResults = inspection.GetInspectionResults().ToList();
                    break;
            }

            if (!inspectionResults.Any()) return null;

            var filteredResults = inspectionResults.GroupBy(x => x.Description).Select(x => x.First()).ToList();
          
            return filteredResults.Select(item => new Issue
                {
                    Severity = item.Inspection.Severity.ToString(),
                    Description = item.Description,
                    Type = item.Inspection.AnnotationName,
                    Meta = item.Inspection.Meta,
                    Name = ExtractIdentifierName(item.Description),
                    Line = item.QualifiedSelection.Selection.StartLine,
                    Column = item.QualifiedSelection.Selection.StartColumn
                })
                .ToList();
        }

        private static string ExtractIdentifierName(string text)
        {
            return text.Contains("Option Explicit") ? "Option Explicit" : text.Substring(text.IndexOf("'", StringComparison.Ordinal), text.LastIndexOf("'", StringComparison.Ordinal) - text.IndexOf("'", StringComparison.Ordinal));
        }

    }
}
