using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog
{
    internal class VbComponents : IVBComponents
    {
        private readonly List<IVBComponent> _components = new List<IVBComponent>();

        public VbComponents(IVBE vbe, IVBProject project)
        {
            VBE = vbe;
            Parent = project;
        }

        public int Count => _components.Count;

        IVBComponent IComCollection<IVBComponent>.this[object index]
        {
            get { throw new NotImplementedException(); }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerator<IVBComponent> GetEnumerator()
        {
            return _components.GetEnumerator();
        }

        public bool Equals(IVBComponents other)
        {
            throw new NotImplementedException();
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }
        IVBComponent IVBComponents.this[object index] => index is string ? _components.Single(x => x.Name == (string)index) : _components.ElementAt((int)index);
        public IVBE VBE { get; }
        public IVBProject Parent { get; }

        public void Remove(IVBComponent item)
        {
            _components.Remove(_components.First(m => m == item));
        }

        public IVBComponent Add(ComponentType type)
        {
            return AddInternal(type, "test");
        }

        public IVBComponent Import(string path)
        {
            ComponentType type = GetComponentTypeFromExtension(path.Split('.').Last());
            string name = path.Split('\\').Last();
            return AddInternal(type, name);
        }

        public IVBComponent AddCustom(string progId)
        {
            throw new NotImplementedException();
        }

        public IVBComponent AddMTDesigner(int index = 0)
        {
            throw new NotImplementedException();
        }

        public void ImportSourceFile(string path)
        {
            throw new NotImplementedException();
        }

        public void RemoveSafely(IVBComponent component)
        {
            throw new NotImplementedException();
        }

        public IVBComponent Add(IVBComponent component)
        {
            _components.Add(component);
            return component;
        }

        private static ComponentType GetComponentTypeFromExtension(string extension)
        {
            ComponentType type;
            new Dictionary<string, ComponentType> {{"bas", ComponentType.StandardModule}, {"cls", ComponentType.ClassModule}, {"frm", ComponentType.UserForm}}
                .TryGetValue(extension, out type);
            return type;
        }

        private VbComponent AddInternal(ComponentType type, string name)
        {
            var component = new VbComponent(VBE, name, type, this);
            var codePane = new CodePane(VBE, new Window(name), new Selection(), component);
            component.CodeModule = codePane.CodeModule = new CodeModule(VBE, name, "", component, codePane);
            _components.Add(component);
            return component;
        }
    }
}