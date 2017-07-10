using System;
using System.Collections;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog
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