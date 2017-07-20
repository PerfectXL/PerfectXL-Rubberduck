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
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
{
    internal class CodePane : ICodePane
    {
        private readonly IVBComponent _vbComponent;

        public CodePane(IVBE vbe, IWindow window, Selection selection, IVBComponent vbComponent)
        {
            VBE = vbe;
            _vbComponent = vbComponent;
            Window = window;
            Selection = selection;

            ((Windows)VBE.Windows).WindowList.Add(Window);
        }

        public IVBE VBE { get; }
        public ICodePanes Collection { get; }
        public IWindow Window { get; }
        public int TopLine { get; set; }
        public int CountOfVisibleLines { get; }
        public ICodeModule CodeModule { get; set; }
        public CodePaneView CodePaneView { get; }
        public Selection Selection { get; set; }

        public QualifiedSelection? GetQualifiedSelection()
        {
            if (Selection.IsEmpty())
            {
                return null;
            }
            return new QualifiedSelection(new QualifiedModuleName(_vbComponent), Selection);
        }

        public void Show()
        {
            throw new NotImplementedException();
        }

        public bool Equals(ICodePane other)
        {
            throw new NotImplementedException();
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }
    }
}
