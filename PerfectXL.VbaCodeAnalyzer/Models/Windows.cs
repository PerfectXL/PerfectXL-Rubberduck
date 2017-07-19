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

using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
{
    internal class Windows : IWindows
    {
        public Windows(IVBE vbe)
        {
            VBE = vbe;
        }

        public IList<IWindow> WindowList { get; } = new List<IWindow>();

        public int Count => WindowList.Count;

        public IWindow this[object index] => index is string ? WindowList.SingleOrDefault(window => window.Caption == (string)index) : WindowList[(int)index];

        IEnumerator IEnumerable.GetEnumerator()
        {
            return WindowList.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return WindowList.GetEnumerator();
        }

        public bool Equals(IWindows other)
        {
            return Equals(this, other);
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }

        public IVBE VBE { get; set; }

        public IApplication Parent => null;

        public ToolWindowInfo CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition)
        {
            return new ToolWindowInfo(new Window(), null);
        }

        public override int GetHashCode()
        {
            return WindowList.GetHashCode();
        }
    }
}
