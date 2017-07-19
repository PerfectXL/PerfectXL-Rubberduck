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
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
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
    }
}
