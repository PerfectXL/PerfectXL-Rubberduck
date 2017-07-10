using System;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog
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

            ((Windows)VBE.Windows).Add(Window);
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