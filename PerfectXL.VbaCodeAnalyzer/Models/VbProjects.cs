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
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
{
    internal class VbProjects : IVBProjects
    {
        public VbProjects(IVBE vbe)
        {
            VBE = vbe;
        }

        public List<IVBProject> Projects { get; } = new List<IVBProject>();

        public int Count => Projects.Count;

        public IVBProject this[object index] => Projects[(int)index];

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerator<IVBProject> GetEnumerator()
        {
            return Projects.GetEnumerator();
        }

        public bool Equals(IVBProjects other)
        {
            throw new NotImplementedException();
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }

        public IVBE VBE { get; }
        public IVBE Parent { get; }

        public IVBProject Add(ProjectType type)
        {
            throw new NotImplementedException();
        }

        public IVBProject Open(string path)
        {
            throw new NotImplementedException();
        }

        public void Remove(IVBProject project)
        {
            throw new NotImplementedException();
        }
    }
}
