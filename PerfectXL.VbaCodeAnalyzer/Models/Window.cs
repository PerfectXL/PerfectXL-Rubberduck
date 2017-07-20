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
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.MSForms;

namespace PerfectXL.VbaCodeAnalyzer.Models
{
    internal class Window : IWindow
    {
        public Window() { }

        public Window(int hWnd)
        {
            HWnd = hWnd;
        }

        public Window(string caption)
        {
            Caption = caption;
        }

        public bool Equals(IWindow other)
        {
            throw new NotImplementedException();
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }

        public int HWnd { get; }
        public string Caption { get; }
        public bool IsVisible { get; set; }
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public WindowState WindowState { get; }
        public WindowKind Type { get; }
        public IVBE VBE { get; }
        public IWindow LinkedWindowFrame { get; }
        public IWindows Collection { get; }
        public ILinkedWindows LinkedWindows { get; }

        public IntPtr Handle()
        {
            throw new NotImplementedException();
        }

        public void Close()
        {
            throw new NotImplementedException();
        }

        public void SetFocus()
        {
            throw new NotImplementedException();
        }

        public void SetKind(WindowKind eKind)
        {
            throw new NotImplementedException();
        }

        public void Detach()
        {
            throw new NotImplementedException();
        }

        public void Attach(int lWindowHandle)
        {
            throw new NotImplementedException();
        }
    }
}
