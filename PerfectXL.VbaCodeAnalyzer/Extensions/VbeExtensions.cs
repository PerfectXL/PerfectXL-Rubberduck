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

using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using PerfectXL.VbaCodeAnalyzer.Inspection;
using PerfectXL.VbaCodeAnalyzer.Models;
using Rubberduck.Common;
using Rubberduck.Parsing.PreProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.DeclarationLoaders;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using AttributeParser = PerfectXL.VbaCodeAnalyzer.Models.AttributeParser;

namespace PerfectXL.VbaCodeAnalyzer.Extensions
{
    internal static class VbeExtensions
    {
        public static ConfiguredParserResult CreateConfiguredParser(this IVBE vbe, string serializedDeclarationsPath = null)
        {
            var state = new RubberduckParserState(vbe, new ConcurrentlyConstructedDeclarationFinderFactory());

            var moduleToModuleReferenceManager = new ModuleToModuleReferenceManager();
            var parserStateManager = new SynchronousParserStateManager(state);
            var referenceRemover = new SynchronousReferenceRemover(state, moduleToModuleReferenceManager);
            var supertypeClearer = new SupertypeClearer(state);
            var comSynchronizer = new SynchronousCOMReferenceSynchronizer(state, parserStateManager, @"C:\");

            var parseRunner = new SynchronousParseRunner(state,
                parserStateManager,
                () => new VBAPreprocessor(double.Parse(vbe.Version, CultureInfo.InvariantCulture)),
                new AttributeParser(),
                new ModuleExporter());

            var declarationResolveRunner = new SynchronousDeclarationResolveRunner(state, parserStateManager, comSynchronizer);
            var referenceResolveRunner = new SynchronousReferenceResolveRunner(state, parserStateManager, moduleToModuleReferenceManager, referenceRemover);

            var parsingCacheService = new ParsingCacheService(state, moduleToModuleReferenceManager, referenceRemover, supertypeClearer);

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

            IProjectManager projectManager = new SynchronousProjectManager(state, vbe);
            var parser = new ParseCoordinator(state, parsingStageService, parsingCacheService, projectManager, parserStateManager, true);

            return new ConfiguredParserResult {ParseCoordinator = parser, ProjectManager = projectManager};
        }

        public static void AddProjectFromCode(this IVBE vbe, string moduleName, string inputCode)
        {
            var project = new VbProject(vbe, "TestProject1", "", ProjectProtection.Unprotected);

            ComponentType componentType = inputCode.GetModuleType();
            string cleanedCodeContent = inputCode.StripVbAttributes();

            project.AddComponent(moduleName, componentType, cleanedCodeContent);
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