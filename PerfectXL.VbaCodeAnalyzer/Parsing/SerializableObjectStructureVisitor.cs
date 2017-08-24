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

using Antlr4.Runtime.Tree;

namespace PerfectXL.VbaCodeAnalyzer.Parsing
{
    internal class SerializableObjectStructureVisitor : IParseTreeVisitor<VbaParseTree>
    {
        public VbaParseTree Visit(IParseTree tree)
        {
            return tree.Accept(this);
        }

        public VbaParseTree VisitChildren(IRuleNode node)
        {
            VbaParseTree result = VbaParseTree.Create(node);

            int childCount = node.ChildCount;
            for (var i = 0; i < childCount; i++)
            {
                VbaParseTree nextResult = node.GetChild(i).Accept(this);
                if (IsNodeToIgnore(nextResult))
                {
                    continue;
                }
                nextResult.Parent = result;
                result.Children.Add(nextResult);
            }
            return result;
        }

        public VbaParseTree VisitTerminal(ITerminalNode node)
        {
            return VbaParseTree.Create(node);
        }

        public VbaParseTree VisitErrorNode(IErrorNode node)
        {
            return new ErrorNode(node);
        }

        private static bool IsNodeToIgnore(VbaParseTree nextResult)
        {
            return string.IsNullOrEmpty(nextResult.Text) && nextResult.ChildCount == 0;
        }
    }
}