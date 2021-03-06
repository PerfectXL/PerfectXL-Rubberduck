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
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
{
    internal class Vbe : IVBE
    {
        private const string VbeVersion = "7.1";

        public Vbe()
        {
            Windows = new Windows(this);
            MainWindow = new Window(0);
            VBProjects = new VbProjects(this);
            CodePanes = new CodePanes(this);
            Version = VbeVersion;
        }

        public bool Equals(IVBE other)
        {
            throw new NotImplementedException();
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }

        public string Version { get; }
        public object HardReference { get; }
        public IWindow ActiveWindow => ActiveCodePane.Window;
        public ICodePane ActiveCodePane { get; set; }
        public IVBProject ActiveVBProject { get; set; }
        public IVBComponent SelectedVBComponent => ActiveCodePane.CodeModule.Parent;
        public IWindow MainWindow { get; }
        public IAddIns AddIns { get; }
        public IVBProjects VBProjects { get; }
        public ICodePanes CodePanes { get; }
        public ICommandBars CommandBars { get; }
        public IWindows Windows { get; }

        public IHostApplication HostApplication()
        {
            return null;
        }

        public IWindow ActiveMDIChild()
        {
            throw new NotImplementedException();
        }

        public bool IsInDesignMode { get; }
    }
}
