using System;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog {
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