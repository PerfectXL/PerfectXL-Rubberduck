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