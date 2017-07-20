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
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
{
    internal class VbComponent : IVBComponent
    {
        public VbComponent(IVBE vbe, string name, ComponentType type, IVBComponents collection)
        {
            VBE = vbe;
            Type = type;
            Collection = collection;
            Name = name;
        }

        public bool Equals(IVBComponent other)
        {
            throw new NotImplementedException();
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }

        public ComponentType Type { get; }
        public ICodeModule CodeModule { get; set; }
        public IVBE VBE { get; }
        public IVBComponents Collection { get; }
        public IProperties Properties { get; }
        public IControls Controls { get; }
        public IControls SelectedControls { get; }
        public bool IsSaved { get; }
        public bool HasDesigner { get; }
        public bool HasOpenDesigner { get; }
        public string DesignerId { get; }
        public string Name { get; set; }

        public IWindow DesignerWindow()
        {
            throw new NotImplementedException();
        }

        public void Activate() { }

        public void Export(string path)
        {
            throw new NotImplementedException();
        }

        public string ExportAsSourceFile(string folder, bool tempFile = false)
        {
            throw new NotImplementedException();
        }

        public IVBProject ParentProject { get; }
    }
}
