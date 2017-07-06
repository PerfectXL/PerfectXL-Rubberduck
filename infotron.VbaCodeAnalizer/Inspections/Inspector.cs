using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using infotron.VbaCodeAnalizer.Mog;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;


namespace infotron.VbaCodeAnalizer.Inspections
{
    public static class Inspector
    {
        public static List<Issue> Inspect<TInspection>(string inputcode, Helper helper) where TInspection : InspectionBase
        {
            IEnumerable<IInspectionResult> inspectionResults = null;

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out component);
            var state = MockParser.CreateAndParse(vbe);

            var inspection = InspectionFactory.Create<TInspection>(state);

            switch (helper)
            {
                case Helper.Yes:
                    var inspector = InspectionsHelper.GetInspector(inspection);
                    inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                    break;
                case Helper.No:
                    inspectionResults = inspection.GetInspectionResults().ToList();
                    break;
            }

            if (!inspectionResults.Any()) return null;

            var filteredResults = inspectionResults.GroupBy(x => x.Description).Select(x => x.First()).ToList();
          
            return filteredResults.Select(item => new Issue
                {
                    Severity = item.Inspection.Severity.ToString(),
                    Description = item.Description,
                    Type = item.Inspection.AnnotationName,
                    Meta = item.Inspection.Meta,
                    Name = ExtractIdentifierName(item.Description),
                    Line = item.QualifiedSelection.Selection.StartLine,
                    Column = item.QualifiedSelection.Selection.StartColumn
                })
                .ToList();
        }

        private static string ExtractIdentifierName(string text)
        {
            return text.Contains("Option Explicit") ? "Option Explicit" : text.Substring(text.IndexOf("'", StringComparison.Ordinal), text.LastIndexOf("'", StringComparison.Ordinal) - text.IndexOf("'", StringComparison.Ordinal));
        }

        public static class InspectionFactory
        {
            public static IInspection Create<TInspection>(RubberduckParserState state) where TInspection : IInspection
            {
                switch (typeof(TInspection).Name)
                {
                    case "ApplicationWorksheetFunctionInspection": return new ApplicationWorksheetFunctionInspection(state); 
                    case "AssignedByValParameterInspection": return new AssignedByValParameterInspection(state);
                    case "ConstantNotUsedInspection": return new ConstantNotUsedInspection(state);
                    case "DefaultProjectNameInspection": return new DefaultProjectNameInspection(state);
                    //case "EmptyIfBlockInspection": return new EmptyIfBlockInspection(state);
                    case "EmptyStringLiteralInspection": return new EmptyStringLiteralInspection(state);
                    case "EncapsulatePublicFieldInspection": return new EncapsulatePublicFieldInspection(state);
                    case "FunctionReturnValueNotUsedInspection": return new FunctionReturnValueNotUsedInspection(state);
                    case "HostSpecificExpressionInspection": return new HostSpecificExpressionInspection(state);
                    // case "HungarianNotationInspection": return new HungarianNotationInspection(state);
                    case "ImplicitActiveSheetReferenceInspection": return new ImplicitActiveSheetReferenceInspection(state);
                    case "ImplicitActiveWorkbookReferenceInspection": return new ImplicitActiveWorkbookReferenceInspection(state);
                    case "ImplicitByRefParameterInspection": return new ImplicitByRefParameterInspection(state);
                    case "ImplicitDefaultMemberAssignmentInspection": return new ImplicitDefaultMemberAssignmentInspection(state);
                    case "ImplicitPublicMemberInspection": return new ImplicitPublicMemberInspection(state);
                    case "ImplicitVariantReturnTypeInspection": return new ImplicitVariantReturnTypeInspection(state);
                    case "MemberNotOnInterfaceInspection": return new MemberNotOnInterfaceInspection(state);
                    case "MissingAnnotationArgumentInspection": return new MissingAnnotationArgumentInspection(state);
                    case "ModuleScopeDimKeywordInspection": return new ModuleScopeDimKeywordInspection(state);
                    case "MoveFieldCloserToUsageInspection": return new MoveFieldCloserToUsageInspection(state);
                    case "MultilineParameterInspection": return new MultilineParameterInspection(state);
                    case "MultipleDeclarationsInspection": return new MultipleDeclarationsInspection(state);
                    case "MultipleFolderAnnotationsInspection": return new MultipleFolderAnnotationsInspection(state);
                    case "NonReturningFunctionInspection": return new NonReturningFunctionInspection(state);
                    case "ObjectVariableNotSetInspection": return new ObjectVariableNotSetInspection(state);
                    case "ObsoleteCallStatementInspection": return new ObsoleteCallStatementInspection(state);
                    case "ObsoleteCommentSyntaxInspection": return new ObsoleteCommentSyntaxInspection(state);
                    case "ObsoleteGlobalInspection": return new ObsoleteGlobalInspection(state);
                    case "ObsoleteLetStatementInspection": return new ObsoleteLetStatementInspection(state);
                    case "ObsoleteTypeHintInspection": return new ObsoleteTypeHintInspection(state);
                    case "OptionBaseInspection": return new OptionBaseInspection(state);
                    case "OptionBaseZeroInspection": return new OptionBaseZeroInspection(state);
                    case "OptionExplicitInspection": return new OptionExplicitInspection(state);
                    case "ParameterCanBeByValInspection": return new ParameterCanBeByValInspection(state);
                    case "ParameterNotUsedInspection": return new ParameterNotUsedInspection(state);
                    case "ProcedureCanBeWrittenAsFunctionInspection": return new ProcedureCanBeWrittenAsFunctionInspection(state);
                    case "ProcedureNotUsedInspection": return new ProcedureNotUsedInspection(state);
                    case "SelfAssignedDeclarationInspection": return new SelfAssignedDeclarationInspection(state);
                    case "UnassignedVariableUsageInspection": return new UnassignedVariableUsageInspection(state);
                    case "UndeclaredVariableInspection": return new UndeclaredVariableInspection(state);
                    case "UntypedFunctionUsageInspection": return new UntypedFunctionUsageInspection(state);
                    // case "UseMeaningfulNameInspection": return new UseMeaningfulNameInspection(state);
                    case "VariableNotAssignedInspection": return new VariableNotAssignedInspection(state);
                    case "VariableNotUsedInspection": return new VariableNotUsedInspection(state);
                    case "VariableTypeNotDeclaredInspection": return new VariableTypeNotDeclaredInspection(state);
                    case "WriteOnlyPropertyInspection": return new WriteOnlyPropertyInspection(state);

                    default: throw new ArgumentException();
                }
            }
        }
    }

}
