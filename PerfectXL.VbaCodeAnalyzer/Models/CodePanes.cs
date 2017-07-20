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
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
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
