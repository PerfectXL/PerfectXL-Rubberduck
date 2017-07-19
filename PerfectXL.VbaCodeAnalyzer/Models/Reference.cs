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
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models {
    internal class Reference : IReference
    {
        public Reference(IVBE vbe, string name, string fullPath, int major, int minor, bool isBuiltIn = true)
        {
            VBE = vbe;
            Name = name;
            FullPath = fullPath;
            Major = major;
            Minor = minor;
            IsBuiltIn = isBuiltIn;
        }

        public bool Equals(IReference other)
        {
            throw new NotImplementedException();
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }

        public string Name { get; }
        public string Guid { get; }
        public string Description { get; }
        public int Major { get; }
        public int Minor { get; }
        public string Version { get; }
        public string FullPath { get; }
        public bool IsBuiltIn { get; }
        public bool IsBroken { get; }
        public ReferenceKind Type { get; }
        public IReferences Collection { get; }
        public IVBE VBE { get; }
    }
}
