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

using System;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace PerfectXL.VbaCodeAnalyzer.Inspection
{
    internal static class InspectionFactory
    {
        public static InspectionWrapper Create<TInspection>(RubberduckParserState state, ResultFetchMethod resultFetchMethod) where TInspection : IInspection
        {
            IInspection inspection = Create<TInspection>(state);
            return new InspectionWrapper(inspection, state, resultFetchMethod);
        }

        private static IInspection Create<TInspection>(RubberduckParserState state) where TInspection : IInspection
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
                //case "HungarianNotationInspection": return new HungarianNotationInspection(state);
                case "ImplicitActiveSheetReferenceInspection": return new ImplicitActiveSheetReferenceInspection(state);
                case "ImplicitActiveWorkbookReferenceInspection": return new ImplicitActiveWorkbookReferenceInspection(state);
                case "ImplicitByRefModifierInspection": return new ImplicitByRefModifierInspection(state);
                case "ImplicitDefaultMemberAssignmentInspection": return new ImplicitDefaultMemberAssignmentInspection(state);
                case "ImplicitPublicMemberInspection": return new ImplicitPublicMemberInspection(state);
                case "ImplicitVariantReturnTypeInspection": return new ImplicitVariantReturnTypeInspection(state);
                case "MemberNotOnInterfaceInspection": return new MemberNotOnInterfaceInspection(state);
                case "MissingAnnotationArgumentInspection": return new MissingAnnotationArgumentInspection(state);
                case "ModuleScopeDimKeywordInspection": return new ModuleScopeDimKeywordInspection(state);
                case "MoveFieldCloserToUsageInspection": return new MoveFieldCloserToUsageInspection(state);
                case "MultilineParameterInspection": return new MultilineParameterInspection(state);
                case "MultipleDeclarationsInspection": return new MultipleDeclarationsInspection(state);
                //case "MultipleFolderAnnotationsInspection": return new MultipleFolderAnnotationsInspection(state);
                case "NonReturningFunctionInspection": return new NonReturningFunctionInspection(state);
                case "ObjectVariableNotSetInspection": return new ObjectVariableNotSetInspection(state);
                case "ObsoleteCallStatementInspection": return new ObsoleteCallStatementInspection(state);
                case "ObsoleteCommentSyntaxInspection": return new ObsoleteCommentSyntaxInspection(state);
                case "ObsoleteGlobalInspection": return new ObsoleteGlobalInspection(state);
                case "ObsoleteLetStatementInspection": return new ObsoleteLetStatementInspection(state);
                case "ObsoleteTypeHintInspection": return new ObsoleteTypeHintInspection(state);
                case "OptionBaseInspection": return new OptionBaseInspection(state);
                case "OptionExplicitInspection": return new OptionExplicitInspection(state);
                case "ParameterCanBeByValInspection": return new ParameterCanBeByValInspection(state);
                case "ParameterNotUsedInspection": return new ParameterNotUsedInspection(state);
                case "ProcedureCanBeWrittenAsFunctionInspection": return new ProcedureCanBeWrittenAsFunctionInspection(state);
                case "ProcedureNotUsedInspection": return new ProcedureNotUsedInspection(state);
                case "SelfAssignedDeclarationInspection": return new SelfAssignedDeclarationInspection(state);
                case "UnassignedVariableUsageInspection": return new UnassignedVariableUsageInspection(state);
                case "UndeclaredVariableInspection": return new UndeclaredVariableInspection(state);
                case "UntypedFunctionUsageInspection": return new UntypedFunctionUsageInspection(state);
                //case "UseMeaningfulNameInspection": return new UseMeaningfulNameInspection(state);
                case "VariableNotAssignedInspection": return new VariableNotAssignedInspection(state);
                case "VariableNotUsedInspection": return new VariableNotUsedInspection(state);
                case "VariableTypeNotDeclaredInspection": return new VariableTypeNotDeclaredInspection(state);
                case "WriteOnlyPropertyInspection": return new WriteOnlyPropertyInspection(state);
                default: throw new ArgumentOutOfRangeException(nameof(TInspection), typeof(TInspection).Name, null);
            }
        }
    }
}
