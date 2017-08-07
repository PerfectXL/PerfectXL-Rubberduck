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

using System.Collections.Generic;
using PerfectXL.VbaCodeAnalyzer.Macro;

namespace PerfectXL.VbaCodeAnalyzer.Inspection
{
    public class CodeInspectionResult
    {
        public CodeInspectionResult(string moduleName)
        {
            ModuleName = moduleName;
        }

        public string ModuleName { get; }
        public List<VbaCodeIssue> VbaCodeIssues { get; set; } = new List<VbaCodeIssue>();
        public List<MacroState> MacroStates { get; set; } = new List<MacroState>();
    }
}
