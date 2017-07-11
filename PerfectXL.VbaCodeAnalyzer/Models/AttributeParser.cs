using System;
using System.Collections.Generic;
using System.Threading;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
{
    internal class AttributeParser : IAttributeParser
    {
        public IDictionary<Tuple<string, DeclarationType>, Attributes> Parse(IVBComponent component, CancellationToken token, out ITokenStream stream,
            out IParseTree tree)
        {
            stream = null;
            tree = null;
            return new Dictionary<Tuple<string, DeclarationType>, Attributes>();
        }
    }
}