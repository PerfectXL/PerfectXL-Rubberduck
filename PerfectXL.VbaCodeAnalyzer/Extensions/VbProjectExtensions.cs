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

using PerfectXL.VbaCodeAnalyzer.Models;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Extensions
{
    internal static class VbProjectExtensions
    {
        public static void AddModuleFromCode(this IVBProject vbProject, string moduleName, string inputCode)
        {
            ComponentType componentType = inputCode.GetModuleType();
            string cleanedCodeContent = inputCode.StripVbAttributes();

            vbProject.AddComponent(moduleName, componentType, cleanedCodeContent);
        }

        private static void AddComponent(this IVBProject vbProject, string name, ComponentType type, string content, Selection selection = new Selection())
        {
            var component = new VbComponent(vbProject.VBE, name, type, vbProject.VBComponents);
            var codePane = new CodePane(vbProject.VBE, new Window(name), selection, component);
            vbProject.VBE.ActiveCodePane = codePane;
            component.CodeModule = codePane.CodeModule = new CodeModule(vbProject.VBE, name, content, component, codePane);
            ((VbComponents)vbProject.VBComponents).Add(component);
        }
    }
}