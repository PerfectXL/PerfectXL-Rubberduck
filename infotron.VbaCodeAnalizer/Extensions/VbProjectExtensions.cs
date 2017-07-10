using PerfectXL.VbaCodeAnalyzer.Models;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Extensions
{
    internal static class VbProjectExtensions
    {
        public static void AddComponent(this IVBProject vbProject, string name, ComponentType type, string content, Selection selection = new Selection())
        {
            var component = new VbComponent(vbProject.VBE, name, type, vbProject.VBComponents);
            var codePane = new CodePane(vbProject.VBE, new Window(name), selection, component);
            vbProject.VBE.ActiveCodePane = codePane;
            component.CodeModule = codePane.CodeModule = new CodeModule(vbProject.VBE, name, content, component, codePane);
            ((VbComponents)vbProject.VBComponents).Add(component);
        }
    }
}