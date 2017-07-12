using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PerfectXL.VbaCodeAnalyzer.Extensions;
using PerfectXL.VbaCodeAnalyzer.Inspection;
using PerfectXL.VbaCodeAnalyzer.Models;
using Rubberduck.Inspections.Abstract;
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
        public IList<CodeInspectionResult> Run(IDictionary<string, string> modules)
        {
            return modules.Select(module => AnalyzeModule(module.Key, module.Value)).ToList();
        }

        internal CodeInspectionResult AnalyzeModule(string moduleName, string moduleCode)
        {
            List<VbaCodeIssue> vbaCodeIssues = new[]
            {
                Inspect<ApplicationWorksheetFunctionInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<AssignedByValParameterInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ConstantNotUsedInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                //Inspect<EmptyIfBlockInspection>(moduleName, moduleCode, ResultFetchMethod.UsingHelper),
                Inspect<EmptyStringLiteralInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<EncapsulatePublicFieldInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<FunctionReturnValueNotUsedInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ImplicitActiveSheetReferenceInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ImplicitActiveWorkbookReferenceInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ImplicitByRefModifierInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ImplicitPublicMemberInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ImplicitVariantReturnTypeInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<MemberNotOnInterfaceInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<MissingAnnotationArgumentInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ModuleScopeDimKeywordInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<MoveFieldCloserToUsageInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<MultilineParameterInspection>(moduleName, moduleCode, ResultFetchMethod.UsingHelper),
                Inspect<MultipleDeclarationsInspection>(moduleName, moduleCode, ResultFetchMethod.UsingHelper),
                Inspect<NonReturningFunctionInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ObsoleteCallStatementInspection>(moduleName, moduleCode, ResultFetchMethod.UsingHelper),
                Inspect<ObsoleteCommentSyntaxInspection>(moduleName, moduleCode, ResultFetchMethod.UsingHelper),
                Inspect<ObsoleteGlobalInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ObsoleteLetStatementInspection>(moduleName, moduleCode, ResultFetchMethod.UsingHelper),
                Inspect<ObsoleteTypeHintInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<OptionBaseInspection>(moduleName, moduleCode, ResultFetchMethod.UsingHelper),
                Inspect<OptionBaseInspection>(moduleName, moduleCode, ResultFetchMethod.UsingHelper),
                Inspect<OptionExplicitInspection>(moduleName, moduleCode, ResultFetchMethod.UsingHelper),
                Inspect<ParameterCanBeByValInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ParameterNotUsedInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<ProcedureCanBeWrittenAsFunctionInspection>(moduleName, moduleCode, ResultFetchMethod.UsingHelper),
                Inspect<ProcedureNotUsedInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<SelfAssignedDeclarationInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<UnassignedVariableUsageInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<UndeclaredVariableInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<UntypedFunctionUsageInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<VariableNotAssignedInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<VariableNotUsedInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<VariableTypeNotDeclaredInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper),
                Inspect<WriteOnlyPropertyInspection>(moduleName, moduleCode, ResultFetchMethod.NoHelper)            }.SelectMany(x => x).ToList();

            return new CodeInspectionResult(moduleName) {VbaCodeIssues = vbaCodeIssues};
        }

        private static string CleanupFileName(string fileName)
        {
            int afterLastHyphenPosition = fileName.LastIndexOf('-') + 1;
            return fileName.Substring(afterLastHyphenPosition, fileName.Length - afterLastHyphenPosition);
        }

        private IEnumerable<VbaCodeIssue> Inspect<TInspection>(string moduleName, string inputCode, ResultFetchMethod resultFetchMethod)
            where TInspection : InspectionBase
        {
            RubberduckParserState parserState = Parse(inputCode);

            IEnumerable<IInspectionResult> inspectionResults = InspectionFactory.Create<TInspection>(parserState, resultFetchMethod).GetInspectionResults();

            return inspectionResults.GroupBy(x => x.Description).Select(x => x.First()).Select(item => new VbaCodeIssue(item, _fileName, moduleName));
        }

        private static RubberduckParserState Parse(string inputCode)
        {
            IVBE vbe = new Vbe();
            vbe.AddProjectFromCode(inputCode);
            ParseCoordinator parser = vbe.CreateConfiguredParser();
            parser.Parse(new CancellationTokenSource());

            return parser.State;
        }
    }
}