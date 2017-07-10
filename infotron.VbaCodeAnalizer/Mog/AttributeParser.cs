using System;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog
{
    internal class AttributeParser : IAttributeParser
    {
        public IDictionary<Tuple<string, DeclarationType>, Attributes> Parse(IVBComponent component, CancellationToken token)
        {
            return new Dictionary<Tuple<string, DeclarationType>, Attributes>();
        }
    }
}