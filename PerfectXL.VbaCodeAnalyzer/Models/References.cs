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

#pragma warning disable 67
        public event EventHandler<ReferenceEventArgs> ItemAdded;
        public event EventHandler<ReferenceEventArgs> ItemRemoved;
#pragma warning restore 67
    }
}
