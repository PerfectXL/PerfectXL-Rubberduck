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
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
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

        private IVBComponent AddInternal(ComponentType type, string name)
        {
            var component = new VbComponent(VBE, name, type, this);
            var codePane = new CodePane(VBE, new Window(name), new Selection(), component);
            component.CodeModule = codePane.CodeModule = new CodeModule(VBE, name, "", component, codePane);

            return Add(component);
        }
    }
}
