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
