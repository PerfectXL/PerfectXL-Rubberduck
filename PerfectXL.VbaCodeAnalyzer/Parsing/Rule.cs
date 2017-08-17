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

using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;

namespace PerfectXL.VbaCodeAnalyzer.Parsing
{
    public class Rule : VbaParseTree
    {
        private readonly RuleContext _ruleContext;
        private readonly string[] _ruleNames = VBAParser.ruleNames;

        public Rule(IParseTree node, RuleContext ruleContext) : base(node)
        {
            _ruleContext = ruleContext;
        }

        public string RuleName => RuleIndex >= 0 && RuleIndex < _ruleNames.Length ? _ruleNames[RuleIndex] : "";
        public int Depth => _ruleContext.Depth();
        public bool IsEmpty => _ruleContext.IsEmpty;
        public int RuleIndex => _ruleContext.RuleIndex;
        public Interval SourceInterval => new Interval(_ruleContext.SourceInterval.a, _ruleContext.SourceInterval.b);
    }
}