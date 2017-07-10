using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog
{
    internal class CodePanes : ICodePanes
    {
        public CodePanes(IVBE vbe)
        {
            VBE = vbe;
        }

        public List<ICodePane> Panes { get; } = new List<ICodePane>();

        public IVBE Parent { get; }
        public IVBE VBE { get; }
        public ICodePane Current { get; set; }

        public int Count => Panes.Count;

        public ICodePane this[object index] => Panes[(int)index];

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerator<ICodePane> GetEnumerator()
        {
            return Panes.GetEnumerator();
        }

        public bool Equals(ICodePanes other)
        {
            throw new NotImplementedException();
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }
    }
}