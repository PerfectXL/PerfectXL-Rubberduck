using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
{
    internal class References : IReferences
    {
        private readonly List<IReference> _references = new List<IReference>();

        public References(IVBE vbe, IVBProject project)
        {
            VBE = vbe;
            Parent = project;
        }

        public int Count => _references.Count;
        public IReference this[object index] => _references.ElementAt((int)index - 1);

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerator<IReference> GetEnumerator()
        {
            return _references.GetEnumerator();
        }

        public bool Equals(IReferences other)
        {
            throw new NotImplementedException();
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }
        public event EventHandler<ReferenceEventArgs> ItemAdded;
        public event EventHandler<ReferenceEventArgs> ItemRemoved;
        public IVBE VBE { get; }
        public IVBProject Parent { get; }

        public IReference AddFromGuid(string guid, int major, int minor)
        {
            throw new NotImplementedException();
        }

        public IReference AddFromFile(string path)
        {
            var reference = new Reference(VBE, path, path, 0, 0);
            _references.Add(reference);
            return reference;
        }

        public void Remove(IReference reference)
        {
            _references.Remove(_references.First(m => m == reference));
        }
    }
}