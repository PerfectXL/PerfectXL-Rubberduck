using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.Inspections.Rubberduck.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace infotron.VbaCodeAnalizer.Inspections
{
    internal static class InspectionExtensions
    {
        public static IEnumerable<IInspectionResult> GetInspectionResults(this IInspection inspection, ResultFetchMethod resultFetchMethod,
            ParseCoordinator parser)
        {
            switch (resultFetchMethod)
            {
                case ResultFetchMethod.UsingHelper: return inspection.GetInspectionResults(parser.State);
                case ResultFetchMethod.NoHelper: return inspection.GetInspectionResults();
                default: throw new ArgumentOutOfRangeException(nameof(resultFetchMethod), resultFetchMethod, null);
            }
        }

        public static IEnumerable<IInspectionResult> GetInspectionResults(this IInspection inspection, RubberduckParserState parserState)
        {
            return inspection.GetInspector(inspection).FindIssuesAsync(parserState, CancellationToken.None).Result;
        }

        public static IInspector GetInspector(this IInspection inspection, params IInspection[] otherInspections)
        {
            return new Inspector(new GeneralConfigService(inspection), otherInspections.Union(new[] {inspection}));
        }
    }
}