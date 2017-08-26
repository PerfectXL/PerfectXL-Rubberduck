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
            return modules.Select(module => AnalyzeModule(module.Key, module.Value)).ToList();
        }

        internal CodeAnalyzerResult AnalyzeModule(string moduleName, string moduleCode)
        {
            RubberduckParseResult rubberduckParseResult = Parse(moduleName, moduleCode);
            RubberduckParserState parserState = rubberduckParseResult.ParserState;

            if (parserState.Status != ParserState.Ready)
            {
                return new CodeAnalyzerResult(moduleName) {VbaCodeIssues = GetModuleExceptions(moduleName, parserState)};
            }

            List<VbaCodeIssue> vbaCodeIssues = new[]
            {
                Inspect<ApplicationWorksheetFunctionInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<AssignedByValParameterInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ConstantNotUsedInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                //Inspect<EmptyIfBlockInspection>(moduleName, parserState, ResultFetchMethod.UsingHelper),
                Inspect<EmptyStringLiteralInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<EncapsulatePublicFieldInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<FunctionReturnValueNotUsedInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ImplicitActiveSheetReferenceInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ImplicitActiveWorkbookReferenceInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ImplicitByRefModifierInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ImplicitPublicMemberInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ImplicitVariantReturnTypeInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<MemberNotOnInterfaceInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<MissingAnnotationArgumentInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ModuleScopeDimKeywordInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<MoveFieldCloserToUsageInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<MultilineParameterInspection>(moduleName, parserState, ResultFetchMethod.UsingHelper),
                Inspect<MultipleDeclarationsInspection>(moduleName, parserState, ResultFetchMethod.UsingHelper),
                Inspect<NonReturningFunctionInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ObsoleteCallStatementInspection>(moduleName, parserState, ResultFetchMethod.UsingHelper),
                Inspect<ObsoleteCommentSyntaxInspection>(moduleName, parserState, ResultFetchMethod.UsingHelper),
                Inspect<ObsoleteGlobalInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ObsoleteLetStatementInspection>(moduleName, parserState, ResultFetchMethod.UsingHelper),
                Inspect<ObsoleteTypeHintInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<OptionBaseInspection>(moduleName, parserState, ResultFetchMethod.UsingHelper),
                Inspect<OptionBaseInspection>(moduleName, parserState, ResultFetchMethod.UsingHelper),
                Inspect<OptionExplicitInspection>(moduleName, parserState, ResultFetchMethod.UsingHelper),
                Inspect<ParameterCanBeByValInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ParameterNotUsedInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<ProcedureCanBeWrittenAsFunctionInspection>(moduleName, parserState, ResultFetchMethod.UsingHelper),
                Inspect<ProcedureNotUsedInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<SelfAssignedDeclarationInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<UnassignedVariableUsageInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<UndeclaredVariableInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<UntypedFunctionUsageInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<VariableNotAssignedInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<VariableNotUsedInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<VariableTypeNotDeclaredInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
                Inspect<WriteOnlyPropertyInspection>(moduleName, parserState, ResultFetchMethod.NoHelper)
            }.SelectMany(x => x).ToList();

            IParseTree parseTree = rubberduckParseResult.GetParseTree(moduleName);
            return new CodeAnalyzerResult(moduleName) {VbaCodeIssues = vbaCodeIssues, ParseTree = parseTree.Accept(new SerializableObjectStructureVisitor())};
        }

        internal RubberduckParseResult Parse(string moduleName, string inputCode)
        {
            IVBE vbe = new Vbe();
            vbe.AddProjectFromCode(moduleName, inputCode);
            ConfiguredParserResult configuredParser = vbe.CreateConfiguredParser();
            configuredParser.ParseCoordinator.Parse(new CancellationTokenSource());
            return new RubberduckParseResult {ParserState = configuredParser.ParseCoordinator.State, ProjectManager = configuredParser.ProjectManager};
        }

        private List<VbaCodeIssue> GetModuleExceptions(string moduleName, RubberduckParserState parserState)
        {
            return parserState.ModuleExceptions
                .Select(x => new VbaCodeIssue(x.Item2, _fileName, moduleName)).ToList();
        }

        private IEnumerable<VbaCodeIssue> Inspect<TInspection>(string moduleName, RubberduckParserState parserState, ResultFetchMethod resultFetchMethod)
            where TInspection : IInspection
        {
            IEnumerable<IInspectionResult> inspectionResults = InspectionFactory.Create<TInspection>(parserState, resultFetchMethod).GetInspectionResults();

            return inspectionResults.GroupBy(x => x.Description).Select(x => x.First()).Select(item => new VbaCodeIssue(item, _fileName, moduleName));
        }
    }
}