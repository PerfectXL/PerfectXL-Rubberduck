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
    public class Token : VbaParseTree
    {
        private readonly CommonToken _token;

        public Token(IParseTree node, CommonToken token) : base(node)
        {
            _token = token;
        }

        public int Column => _token.Column;
        public int Line => _token.Line;
        public int StartIndex => _token.StartIndex;
        public int StopIndex => _token.StopIndex;
        public string RuleName => GetRuleName(TokenType);
        public int TokenType => _token.Type;

        private static string GetRuleName(int i)
        {
            int j = i - 1;
            if (j <= VBALexer.FLOATLITERAL && j >= 0 && j < VBALexer.ruleNames.Length)
            {
                return VBALexer.ruleNames[j];
            }
            switch (j)
            {
                case 227: return "NEWLINE";
                case 228: return "SINGLEQUOTE";
                case 229: return "UNDERSCORE";
                case 230: return "WS";
                case 231: return "GUIDLITERAL";
                case 232: return "IDENTIFIER";
                case 233: return "LINE_CONTINUATION";
                case 234: return "ERRORCHAR";
                default: return "(UNKOWN)";
            }
        }
    }
}