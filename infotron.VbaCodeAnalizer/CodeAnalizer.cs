using System;
using System.Collections.Generic;
using System.Linq;
using infotron.VbaCodeAnalizer.Inspections;
using Rubberduck.Inspections.Concrete;

namespace infotron.VbaCodeAnalizer
{
    public static class CodeAnalizer
    {
        /// <summary>
        /// Inspects vba Code and returns code issues,
        /// </summary>
        /// <param name="modules"></param>
        /// <param name="filename"></param>
        /// <returns>Json string</returns>
        public static List<CodeInspection> Run(Dictionary<string, string> modules, string filename)
        {
            return modules.Select(module => Analize(module.Value, module.Key, filename)).ToList();
        }

        private static CodeInspection Analize(string code, string modulename, string documentname)
        {
            var codeinspection = new CodeInspection();

            var strt = documentname.LastIndexOf("-", StringComparison.Ordinal) + 1;
            var file = documentname.Substring(strt, documentname.Length - strt);

            #region EmptyIfBlockInspector
            //var issues = Inspector.Inspect<EmptyIfBlockInspection>(code, Helper.Yes);
            //if (issues != null)
            //{
            //   issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
            //    codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            //}
            #endregion

            #region StringLiteralInspector
            var issues = Inspector.Inspect<EmptyStringLiteralInspection>(code, Helper.No);

            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region AssignedByValParameterInspector
            issues = Inspector.Inspect<AssignedByValParameterInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ConstantNotUsedInspector
            issues = Inspector.Inspect<ConstantNotUsedInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region EncapsulatePublicFieldInspector
            issues = Inspector.Inspect<EncapsulatePublicFieldInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ApplicationWorksheetFunctionInspector
            issues = Inspector.Inspect<ApplicationWorksheetFunctionInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region FunctionReturnValueNotUsedInspector
            issues = Inspector.Inspect<FunctionReturnValueNotUsedInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ImplicitActiveSheetReferenceInspector
            issues = Inspector.Inspect<ImplicitActiveSheetReferenceInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ImplicitActiveWorkbookReferenceInspector
            issues = Inspector.Inspect<ImplicitActiveWorkbookReferenceInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ImplicitByRefParameterInspector
            issues = Inspector.Inspect<ImplicitByRefParameterInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ImplicitPublicMemberInspector
            issues = Inspector.Inspect<ImplicitPublicMemberInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ImplicitVariantReturnTypeInspector
            issues = Inspector.Inspect<ImplicitVariantReturnTypeInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region MemberNotOnInterfaceInspector
            issues = Inspector.Inspect<MemberNotOnInterfaceInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region MissingAnnotationArgumentInspector
            issues = Inspector.Inspect<MissingAnnotationArgumentInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ModuleScopeDimKeywordInspector
            issues = Inspector.Inspect<ModuleScopeDimKeywordInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region MoveFieldCloseToUsageInspector
            issues = Inspector.Inspect<MoveFieldCloserToUsageInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region MultilineParameterInspector
            issues = Inspector.Inspect<MultilineParameterInspection>(code, Helper.Yes);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region MultipleDeclarationsInspector
            issues = Inspector.Inspect<MultipleDeclarationsInspection>(code, Helper.Yes);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region MultipleFolderAnnotationsInspector
            issues = Inspector.Inspect<MultipleFolderAnnotationsInspection>(code, Helper.Yes);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region NonReturningFunctionInspector
            issues = Inspector.Inspect<NonReturningFunctionInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ObsoleteCallStatementInspector
            issues = Inspector.Inspect<ObsoleteCallStatementInspection>(code, Helper.Yes);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ObsoleteCommentSyntaxInspector
            issues = Inspector.Inspect<ObsoleteCommentSyntaxInspection>(code, Helper.Yes);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ObsoleteGlobalInspector
            issues = Inspector.Inspect<ObsoleteGlobalInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ObsoleteLetStatementInspector
            issues = Inspector.Inspect<ObsoleteLetStatementInspection>(code, Helper.Yes);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ObsoleteTypeHintInspector
            issues = Inspector.Inspect<ObsoleteTypeHintInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region OptionBaseInspector
            issues = Inspector.Inspect<OptionBaseInspection>(code, Helper.Yes);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region OptionBaseZeroInspector
            issues = Inspector.Inspect<OptionBaseZeroInspection>(code, Helper.Yes);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region OptionExplicitInspector
            issues = Inspector.Inspect<OptionExplicitInspection>(code, Helper.Yes);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ParameterCanBeByValInspector
            issues = Inspector.Inspect<ParameterCanBeByValInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ParameterNotUsedInspector
            issues = Inspector.Inspect<ParameterNotUsedInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ProcedureNotUsedInspector
            issues = Inspector.Inspect<ProcedureNotUsedInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region ProcedureShouldBeFunctionInspector
            issues = Inspector.Inspect<ProcedureCanBeWrittenAsFunctionInspection>(code, Helper.Yes);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region SelfAssignedDeclarationInspector
            issues = Inspector.Inspect<SelfAssignedDeclarationInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region UnassignedVariableUsageInspector
            issues = Inspector.Inspect<UnassignedVariableUsageInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region UndeclaredVariableInspector
            issues = Inspector.Inspect<UndeclaredVariableInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region UntypedFunctionUsageInspector
            issues = Inspector.Inspect<UntypedFunctionUsageInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region VariableNotAssignedInspector
            issues = Inspector.Inspect<VariableNotAssignedInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region VariableNotUsedInspector
            issues = Inspector.Inspect<VariableNotUsedInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region VariableTypeNotDeclaredInspector
            issues = Inspector.Inspect<VariableTypeNotDeclaredInspection>(code, Helper.No);
            if (issues != null)
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }
            #endregion

            #region WriteOnlyPropertyInspector
            issues = Inspector.Inspect<WriteOnlyPropertyInspection>(code, Helper.No);
            if (issues == null) return codeinspection;
            {
                issues.ForEach(item => { item.ModuleName = modulename; item.FileName = file; });
                codeinspection.MaintabilityAndReadabilityIssues.AddRange(issues);
            }

            #endregion

            return codeinspection;
        }
    }
}
