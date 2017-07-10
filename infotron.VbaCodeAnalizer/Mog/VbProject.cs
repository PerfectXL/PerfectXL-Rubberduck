using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog
{
    internal class VbProject : IVBProject
    {
        public VbProject(IVBE vbe, string name, string fileName, ProjectProtection protection)
        {
            VBE = vbe;
            Name = name;
            FileName = fileName;
            Protection = protection;
            VBComponents = new VbComponents(VBE, this);
        }

        public bool Equals(IVBProject other)
        {
            throw new NotImplementedException();
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }

        public IApplication Application { get; }
        public IApplication Parent { get; }
        public IVBE VBE { get; }
        public IVBProjects Collection { get; }
        public IReferences References => new References(VBE, this);
        public IVBComponents VBComponents { get; }
        public string ProjectId => HelpFile;
        public string Name { get; set; }
        public string Description { get; set; }
        public string HelpFile { get; set; }
        public string FileName { get; }
        public string BuildFileName { get; }
        public bool IsSaved { get; }
        public ProjectType Type { get; }
        public EnvironmentMode Mode { get; }
        public ProjectProtection Protection { get; }

        public void AssignProjectId()
        {
            HelpFile = Guid.NewGuid().ToString();
        }

        public void SaveAs(string fileName)
        {
            throw new NotImplementedException();
        }

        public void MakeCompiledFile()
        {
            throw new NotImplementedException();
        }

        public void ExportSourceFiles(string folder)
        {
            throw new NotImplementedException();
        }

        public string ProjectDisplayName { get; }

        public IReadOnlyList<string> ComponentNames()
        {
            return VBComponents.Select(component => component.Name).ToArray();
        }

        public void AddComponent(string name, ComponentType type, string content, Selection selection = new Selection())
        {
            var component = new VbComponent(VBE, name, type, VBComponents);
            var codePane = new CodePane(VBE, new Window(name), selection, component);
            VBE.ActiveCodePane = codePane;
            component.CodeModule = codePane.CodeModule = new CodeModule(VBE, name, content, component, codePane);
            ((VbComponents)VBComponents).Add(component);
        }
    }
}