using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog
{
    internal static class VbeExtensions
    {
        public static ParseCoordinator CreateConfiguredParser(this IVBE vbe, string serializedDeclarationsPath = null)
        {
            var state = new RubberduckParserState(vbe);

            var moduleToModuleReferenceManager = new ModuleToModuleReferenceManager();
            var parserStateManager = new SynchronousParserStateManager(state);
            var referenceRemover = new SynchronousReferenceRemover(state, moduleToModuleReferenceManager);
            var comSynchronizer = new SynchronousCOMReferenceSynchronizer(state, parserStateManager, @"C:\");

            var parseRunner = new SynchronousParseRunner(state,
                parserStateManager,
                () => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture)),
                new AttributeParser());

            var declarationResolveRunner = new SynchronousDeclarationResolveRunner(state, parserStateManager, comSynchronizer);
            var referenceResolveRunner = new SynchronousReferenceResolveRunner(state, parserStateManager, moduleToModuleReferenceManager, referenceRemover);

            var parsingStageService = new ParsingStageService(comSynchronizer,
                new BuiltInDeclarationLoader(state,
                    new List<ICustomDeclarationLoader>
                    {
                        new DebugDeclarations(state),
                        new SpecialFormDeclarations(state),
                        new FormEventDeclarations(state),
                        new AliasDeclarations(state)
                    }),
                parseRunner,
                declarationResolveRunner,
                referenceResolveRunner);

            var parser = new ParseCoordinator(state,
                parsingStageService,
                new SynchronousProjectManager(state, vbe),
                moduleToModuleReferenceManager,
                parserStateManager,
                referenceRemover,
                true);
            return parser;
        }

        public static void AddProjectFromCode(this IVBE vbe, string inputCode)
        {
            var project = new VbProject(vbe, "TestProject1", "", ProjectProtection.Unprotected);
            project.AddComponent("TestModule1", ComponentType.StandardModule, inputCode);

            vbe.AddProject(project);
        }

        private static void AddProject(this IVBE vbe, IVBProject project)
        {
            ((VbProjects)vbe.VBProjects).Projects.Add(project);
            foreach (IVBComponent component in ((VbProjects)vbe.VBProjects).Projects.SelectMany(x => x.VBComponents))
            {
                ((CodePanes)vbe.CodePanes).Panes.Add(component.CodeModule.CodePane);
            }
            vbe.ActiveVBProject = project;
            vbe.ActiveCodePane = project.VBComponents[0].CodeModule.CodePane;
        }
    }
}