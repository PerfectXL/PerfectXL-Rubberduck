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
using System.Text.RegularExpressions;
using System.Threading;
using PerfectXL.VbaCodeAnalyzer.Extensions;
using PerfectXL.VbaCodeAnalyzer.Inspection;
using PerfectXL.VbaCodeAnalyzer.Macro;
using PerfectXL.VbaCodeAnalyzer.Models;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Configuration;

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
        public IList<CodeInspectionResult> Run(IDictionary<string, string> modules)
        {
            return modules.Select(module => AnalyzeModule(module.Key, module.Value)).ToList();
        }

        internal CodeInspectionResult AnalyzeModule(string moduleName, string moduleCode)
        {
            moduleCode = RemoveAttributeLineFromCode(moduleCode);

            RubberduckParserState parserState = Parse(moduleCode);

          //  var test = IdentiefierFilter(parserState);

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
                Inspect<WriteOnlyPropertyInspection>(moduleName, parserState, ResultFetchMethod.NoHelper),
            }.SelectMany(x => x).ToList();

            var inspectionResult = new CodeInspectionResult(moduleName)
            {
                VbaCodeIssues = IssueFilter(vbaCodeIssues)  
            };

            inspectionResult.VbaCodeIssues.AddRange(RankMacro(moduleName, moduleCode));

            return inspectionResult;
        }


        internal List<VbaCodeIssue> RankMacro(string moduleName, string moduleCode)
        {
            moduleCode = RemoveAttributeLineFromCode(moduleCode);
            return MacroInspector.Run(Parse(moduleCode));
        }

        internal RubberduckParserState Parse(string inputCode)
        {
            IVBE vbe = new Vbe();
            vbe.AddProjectFromCode(inputCode);
            ParseCoordinator parser = vbe.CreateConfiguredParser();
            parser.Parse(new CancellationTokenSource());

            return parser.State;
        }

        internal IVBE GetVbe(string inputCode)
        {
            IVBE vbe = new Vbe();
            vbe.AddProjectFromCode(inputCode);
            return vbe;
        }

        private IEnumerable<VbaCodeIssue> Inspect<TInspection>(string moduleName, RubberduckParserState parserState, ResultFetchMethod resultFetchMethod) where TInspection : IInspection
        {
            IEnumerable<IInspectionResult> inspectionResults = InspectionFactory.Create<TInspection>(parserState, resultFetchMethod).GetInspectionResults();

            return inspectionResults.GroupBy(x => x.Description).Select(x => x.First()).Select(item => new VbaCodeIssue(item, _fileName, moduleName));
        }

        private static string RemoveAttributeLineFromCode(string code)
        {
            return string.Join("\r\n", Regex.Split(code, "\r\n").Where(s => !Regex.IsMatch(s, "A?ttribute VB_")));
        }

        private static List<VbaCodeIssue> IssueFilter(List<VbaCodeIssue> vbaCodeIssues)
        {
            var vbaTerms = ConfigurationManager.AppSettings["VbaTerms"].Split(',');

             var filteredCodeIssues = vbaCodeIssues.FindAll(terms => !vbaTerms.Contains(terms.Name));

            filteredCodeIssues = filteredCodeIssues.FindAll(s => !Regex.IsMatch(s.Name, @"^xl", RegexOptions.IgnoreCase));

            filteredCodeIssues = filteredCodeIssues.FindAll(s => !Regex.IsMatch(s.Type, @"^UnassignedVariableUsage", RegexOptions.IgnoreCase));

            return filteredCodeIssues;
        }
    }
}
