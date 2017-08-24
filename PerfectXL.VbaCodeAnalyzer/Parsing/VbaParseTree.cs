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
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;

namespace PerfectXL.VbaCodeAnalyzer.Parsing
{
    public abstract class VbaParseTree
    {
        private readonly IParseTree _node;

        protected VbaParseTree(IParseTree node)
        {
            _node = node;
        }

        public VbaParseTree Parent { get; set; }
        public int ChildCount => _node.ChildCount;
        public IList<VbaParseTree> Children { get; } = new List<VbaParseTree>();
        public string Text => _node.GetText();

        public static VbaParseTree Create(IParseTree node)
        {
            var ruleContext = node.Payload as RuleContext;
            if (ruleContext != null)
            {
                return new Rule(node, ruleContext);
            }
            var token = node.Payload as CommonToken;
            if (token != null)
            {
                return new Token(node, token);
            }
            return new UnkownNode(node);
        }
    }
}