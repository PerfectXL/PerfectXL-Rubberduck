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

using System.Linq;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace PerfectXL.VbaCodeAnalyzer.Inspection
{
    internal struct RubberduckParseResult
    {
        public IProjectManager ProjectManager { get; set; }
        public RubberduckParserState ParserState { get; set; }

        public IParseTree GetParseTree(string moduleName)
        {
            QualifiedModuleName qualifiedModuleName = ProjectManager.AllModules().FirstOrDefault(x => x.ComponentName == moduleName);
            return qualifiedModuleName != default(QualifiedModuleName) ? ParserState.GetParseTree(qualifiedModuleName) : null;
        }
    }
}