using System;
using System.Collections.Generic;
using System.Threading;
using PerfectXL.VbaCodeAnalyzer.Models;
using Rubberduck.Inspections.Rubberduck.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace PerfectXL.VbaCodeAnalyzer.Inspection
{
    internal class InspectionWrapper
    {
        private readonly IInspection _inspection;
        private readonly RubberduckParserState _parserState;
        private readonly ResultFetchMethod _resultFetchMethod;

        public InspectionWrapper(IInspection inspection, RubberduckParserState parserState, ResultFetchMethod resultFetchMethod)
        {
            _inspection = inspection;
            _parserState = parserState;
            _resultFetchMethod = resultFetchMethod;
        }

        public IEnumerable<IInspectionResult> GetInspectionResults()
        {
            switch (_resultFetchMethod)
            {
                case ResultFetchMethod.UsingHelper: return GetInspectionResults1();
                case ResultFetchMethod.NoHelper: return _inspection.GetInspectionResults();
                default: throw new ArgumentOutOfRangeException(nameof(_resultFetchMethod), _resultFetchMethod, null);
            }
        }

        private IEnumerable<IInspectionResult> GetInspectionResults1()
        {
            return GetInspector().FindIssuesAsync(_parserState, CancellationToken.None).Result;
        }

        private IInspector GetInspector()
        {
            return new Inspector(new GeneralConfigService(_inspection), new[] {_inspection});
        }
    }
}