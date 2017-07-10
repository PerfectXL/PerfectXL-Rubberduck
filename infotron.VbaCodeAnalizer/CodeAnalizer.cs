using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using infotron.VbaCodeAnalizer.Inspections;
using infotron.VbaCodeAnalizer.Mog;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer
{
    public static class CodeAnalizer
    {
        /// <summary>
        ///     Inspects vba Code and returns code issues,
        /// </summary>
        /// <param name="modules"></param>
        /// <param name="filename"></param>
        /// <returns>Json string</returns>
        public static List<CodeInspectionResult> Run(Dictionary<string, string> modules, string filename)
        {
            return modules.Select(module => Analize(module.Value, module.Key, filename)).ToList();
        }

        public static List<VbaCodeIssue> Inspect<TInspection>(string inputCode, ResultFetchMethod resultFetchMethod) where TInspection : InspectionBase
        {
            IVBE vbe = new Vbe();
            vbe.AddProjectFromCode(inputCode);
            ParseCoordinator parser = vbe.CreateConfiguredParser();
            parser.Parse(new CancellationTokenSource());

            IInspection inspection = InspectionFactory.Create<TInspection>(parser.State);

            IEnumerable<IInspectionResult> inspectionResults = inspection.GetInspectionResults(resultFetchMethod, parser);

            return inspectionResults.GroupBy(x => x.Description).Select(x => x.First()).Select(item => new VbaCodeIssue(item)).ToList();
        }

        internal static CodeInspectionResult Analize(string code, string modulename, string documentname)
        {
            var codeinspection = new CodeInspectionResult();

            int strt = documentname.LastIndexOf("-", StringComparison.Ordinal) + 1;
            string file = documentname.Substring(strt, documentname.Length - strt);

            #region EmptyIfBlockInspector
            //var issues = Inspect<EmptyIfBlockInspection>(code, Helper.Yes);
            //if (issues != null)
            //{
            //   issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
            //    codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            //}
            #endregion

            #region StringLiteralInspector
            List<VbaCodeIssue> issues = Inspect<EmptyStringLiteralInspection>(code, ResultFetchMethod.NoHelper);

            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region AssignedByValParameterInspector
            issues = Inspect<AssignedByValParameterInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ConstantNotUsedInspector
            issues = Inspect<ConstantNotUsedInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region EncapsulatePublicFieldInspector
            issues = Inspect<EncapsulatePublicFieldInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ApplicationWorksheetFunctionInspector
            issues = Inspect<ApplicationWorksheetFunctionInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region FunctionReturnValueNotUsedInspector
            issues = Inspect<FunctionReturnValueNotUsedInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ImplicitActiveSheetReferenceInspector
            issues = Inspect<ImplicitActiveSheetReferenceInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ImplicitActiveWorkbookReferenceInspector
            issues = Inspect<ImplicitActiveWorkbookReferenceInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ImplicitByRefParameterInspector
            issues = Inspect<ImplicitByRefParameterInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ImplicitPublicMemberInspector
            issues = Inspect<ImplicitPublicMemberInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ImplicitVariantReturnTypeInspector
            issues = Inspect<ImplicitVariantReturnTypeInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region MemberNotOnInterfaceInspector
            issues = Inspect<MemberNotOnInterfaceInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region MissingAnnotationArgumentInspector
            issues = Inspect<MissingAnnotationArgumentInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ModuleScopeDimKeywordInspector
            issues = Inspect<ModuleScopeDimKeywordInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region MoveFieldCloseToUsageInspector
            issues = Inspect<MoveFieldCloserToUsageInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region MultilineParameterInspector
            issues = Inspect<MultilineParameterInspection>(code, ResultFetchMethod.UsingHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region MultipleDeclarationsInspector
            issues = Inspect<MultipleDeclarationsInspection>(code, ResultFetchMethod.UsingHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region MultipleFolderAnnotationsInspector
            issues = Inspect<MultipleFolderAnnotationsInspection>(code, ResultFetchMethod.UsingHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region NonReturningFunctionInspector
            issues = Inspect<NonReturningFunctionInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ObsoleteCallStatementInspector
            issues = Inspect<ObsoleteCallStatementInspection>(code, ResultFetchMethod.UsingHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ObsoleteCommentSyntaxInspector
            issues = Inspect<ObsoleteCommentSyntaxInspection>(code, ResultFetchMethod.UsingHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ObsoleteGlobalInspector
            issues = Inspect<ObsoleteGlobalInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ObsoleteLetStatementInspector
            issues = Inspect<ObsoleteLetStatementInspection>(code, ResultFetchMethod.UsingHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ObsoleteTypeHintInspector
            issues = Inspect<ObsoleteTypeHintInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region OptionBaseInspector
            issues = Inspect<OptionBaseInspection>(code, ResultFetchMethod.UsingHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region OptionBaseZeroInspector
            issues = Inspect<OptionBaseZeroInspection>(code, ResultFetchMethod.UsingHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region OptionExplicitInspector
            issues = Inspect<OptionExplicitInspection>(code, ResultFetchMethod.UsingHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ParameterCanBeByValInspector
            issues = Inspect<ParameterCanBeByValInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ParameterNotUsedInspector
            issues = Inspect<ParameterNotUsedInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ProcedureNotUsedInspector
            issues = Inspect<ProcedureNotUsedInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region ProcedureShouldBeFunctionInspector
            issues = Inspect<ProcedureCanBeWrittenAsFunctionInspection>(code, ResultFetchMethod.UsingHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region SelfAssignedDeclarationInspector
            issues = Inspect<SelfAssignedDeclarationInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region UnassignedVariableUsageInspector
            issues = Inspect<UnassignedVariableUsageInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region UndeclaredVariableInspector
            issues = Inspect<UndeclaredVariableInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region UntypedFunctionUsageInspector
            issues = Inspect<UntypedFunctionUsageInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region VariableNotAssignedInspector
            issues = Inspect<VariableNotAssignedInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region VariableNotUsedInspector
            issues = Inspect<VariableNotUsedInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region VariableTypeNotDeclaredInspector
            issues = Inspect<VariableTypeNotDeclaredInspection>(code, ResultFetchMethod.NoHelper);
            if (issues != null)
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            #region WriteOnlyPropertyInspector
            issues = Inspect<WriteOnlyPropertyInspection>(code, ResultFetchMethod.NoHelper);
            if (issues == null)
            {
                return codeinspection;
            }
            {
                issues.ForEach(item =>
                {
                    item.ModuleName = modulename;
                    item.FileName = file;
                });
                codeinspection.VbaCodeIssues.AddRange(issues);
            }
            #endregion

            return codeinspection;
        }
    }
}