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
using System.Linq;
using System.Threading;
using Antlr4.Runtime.Tree;
using PerfectXL.VbaCodeAnalyzer.Extensions;
using PerfectXL.VbaCodeAnalyzer.Inspection;
using PerfectXL.VbaCodeAnalyzer.Models;
using PerfectXL.VbaCodeAnalyzer.Parsing;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer
{
    public class CodeAnalyzer
    {
        private readonly string _fileName;

        /// <summary>Constructor</summary>
        /// <param name="fileName">The file name.</param>
        public CodeAnalyzer(string fileName)
        {
            _fileName = fileName;
        }

        /// <summary>
        ///     Inspects VBA code and returns code issues.
        /// </summary>
        /// <param name="modules">A dictionary containing key value pairs of module name and module code (as a string).</param>
        /// <returns>A list of code inspection results per module.</returns>
        public IList<CodeAnalyzerResult> Run(IDictionary<string, string> modules)
        {
            return AnalyzeModules(modules).ToArray();
        }

        /// <remarks>
        ///     For unit testing only.
        /// </remarks>
        internal CodeAnalyzerResult AnalyzeModule(string moduleName, string moduleCode)
        {
            return AnalyzeModules(new Dictionary<string, string> {{moduleName, moduleCode}}).FirstOrDefault();
        }

        private IEnumerable<CodeAnalyzerResult> AnalyzeModules(IDictionary<string, string> modules)
        {
            RubberduckParseResult rubberduckParseResult = Parse(modules);
            RubberduckParserState parserState = rubberduckParseResult.ParserState;

            if (parserState.Status != ParserState.Ready)
            {
                List<VbaCodeIssue> moduleExceptions = GetModuleExceptions(parserState);
                foreach (string moduleName in modules.Keys)
                {
                    VbaCodeIssue[] moduleIssues = moduleExceptions.Where(x => x.ModuleName == moduleName).ToArray();
                    yield return new CodeAnalyzerResult(moduleName) {VbaCodeIssues = moduleIssues};
                }
                yield break;
            }

            VbaCodeIssue[] vbaCodeIssues = new[]
            {
                Inspect<ApplicationWorksheetFunctionInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<AssignedByValParameterInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ConstantNotUsedInspection>(parserState, ResultFetchMethod.NoHelper),
                //Inspect<EmptyIfBlockInspection>(parserState, ResultFetchMethod.UsingHelper),
                Inspect<EmptyStringLiteralInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<EncapsulatePublicFieldInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<FunctionReturnValueNotUsedInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ImplicitActiveSheetReferenceInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ImplicitActiveWorkbookReferenceInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ImplicitByRefModifierInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ImplicitPublicMemberInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ImplicitVariantReturnTypeInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<MemberNotOnInterfaceInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<MissingAnnotationArgumentInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ModuleScopeDimKeywordInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<MoveFieldCloserToUsageInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<MultilineParameterInspection>(parserState, ResultFetchMethod.UsingHelper),
                Inspect<MultipleDeclarationsInspection>(parserState, ResultFetchMethod.UsingHelper),
                Inspect<NonReturningFunctionInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ObsoleteCallStatementInspection>(parserState, ResultFetchMethod.UsingHelper),
                Inspect<ObsoleteCommentSyntaxInspection>(parserState, ResultFetchMethod.UsingHelper),
                Inspect<ObsoleteGlobalInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ObsoleteLetStatementInspection>(parserState, ResultFetchMethod.UsingHelper),
                Inspect<ObsoleteTypeHintInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<OptionBaseInspection>(parserState, ResultFetchMethod.UsingHelper),
                Inspect<OptionExplicitInspection>(parserState, ResultFetchMethod.UsingHelper),
                Inspect<ParameterCanBeByValInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ParameterNotUsedInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<ProcedureCanBeWrittenAsFunctionInspection>(parserState, ResultFetchMethod.UsingHelper),
                Inspect<ProcedureNotUsedInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<SelfAssignedDeclarationInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<UnassignedVariableUsageInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<UndeclaredVariableInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<UntypedFunctionUsageInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<VariableNotAssignedInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<VariableNotUsedInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<VariableTypeNotDeclaredInspection>(parserState, ResultFetchMethod.NoHelper),
                Inspect<WriteOnlyPropertyInspection>(parserState, ResultFetchMethod.NoHelper)
            }.SelectMany(x => x).Select(x => new VbaCodeIssue(x, _fileName)).ToArray();

            foreach (string moduleName in modules.Keys)
            {
                IParseTree parseTree = rubberduckParseResult.GetParseTree(moduleName);
                VbaCodeIssue[] moduleIssues = vbaCodeIssues.Where(x => x.ModuleName == moduleName).ToArray();
                yield return new CodeAnalyzerResult(moduleName) {VbaCodeIssues = moduleIssues, ParseTree = parseTree.Accept(new SerializableObjectStructureVisitor())};
            }
        }

        internal RubberduckParseResult Parse(IDictionary<string, string> modules)
        {
            IVBE vbe = new Vbe();
            IVBProject project = CreateVbProjectWithModules(vbe, modules);
            vbe.AddProject(project);

            ConfiguredParserResult configuredParser = vbe.CreateConfiguredParser();
            configuredParser.ParseCoordinator.Parse(new CancellationTokenSource());
            return new RubberduckParseResult {ParserState = configuredParser.ParseCoordinator.State, ProjectManager = configuredParser.ProjectManager};
        }

        private IVBProject CreateVbProjectWithModules(IVBE vbe, IDictionary<string, string> modules)
        {
            IVBProject project = new VbProject(vbe, "TestProject1", _fileName, ProjectProtection.Unprotected);
            foreach (KeyValuePair<string, string> module in modules)
            {
                project.AddModuleFromCode(module.Key, module.Value);
            }
            return project;
        }

        private List<VbaCodeIssue> GetModuleExceptions(RubberduckParserState parserState)
        {
            return parserState.ModuleExceptions.Select(x => new VbaCodeIssue(x.Item2, _fileName, x.Item1.CodeModule.Name)).ToList();
        }

        private static IEnumerable<IInspectionResult> Inspect<TInspection>(RubberduckParserState parserState, ResultFetchMethod resultFetchMethod)
            where TInspection : IInspection
        {
            IEnumerable<IInspectionResult> inspectionResults = InspectionFactory.Create<TInspection>(parserState, resultFetchMethod).GetInspectionResults();

            return inspectionResults.DistinctBy(x => x.Description);
        }
    }
}
