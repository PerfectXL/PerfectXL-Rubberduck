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

using System.Diagnostics;
using System.Linq;
using NUnit.Framework;
using PerfectXL.VbaCodeAnalyzer.Inspection;

namespace PerfectXL.VbaCodeAnalyzer.UnitTests
{
    [TestFixture]
    public class TypeSpecificTests
    {
        [Test, Ignore]
        public void ApplicationWorksheetFunctionInspectionTest()
        {
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1",
                @"
                Option Explicit

                Private Sub MySub()
                    Dim r As Range, m As Variant
                    Set r = Worksheets(""Sheet1"").Range(""A1:C10"")
                    m = Application.WorksheetFunction.Min(r)
                    MsgBox m
                End Sub
                ");
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ApplicationWorksheetFunction"));
        }

        [Test]
        public void ConstantNotUsedTest()
        {
            const string inputCode = @"Option Explicit
                Public Sub Foo()
                    Const const1 As Integer = 9
                End Sub";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ConstantNotUsed"));
        }

        [Test]
        public void EncapsulatePublicFieldInspectionTest()
        {
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1",
                @"
                Option Explicit

                Public x As New Excel.Application

                Private Sub MySub()
                    x.Calculate
                End Sub
                ");
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "EncapsulatePublicField"));
        }

        [Test]
        public void FunctionNonReturningTest()
        {
            const string inputCode = @"Option Explicit
                    Function Foo(ByVal arg1 As Integer) As Boolean
                        arg1 = 9
                    End Function";
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "NonReturningFunction"));
        }

        [Test]
        public void FunctionReturnValueNotUsedInspectionTest()
        {
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1",
                @"
                Option Explicit

                Private Function MyFunction() As Integer
                    MyFunction = 0
                End Function

                Private Sub MySub()
                    MyFunction
                End Sub
                ");
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "FunctionReturnValueNotUsed"));
        }

        [Test]
        public void FunctionReturnValueNotUsedTest()
        {
            const string inputCode = @"Option Explicit
                    Function Foo(ByVal arg1 As Integer) As Boolean
                        arg1 = 9
                    End Function";
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "FunctionReturnValueNotUsed"));
        }

        [Test]
        public void ImplicitPublicMemberTest()
        {
            const string inputCode = @"option explicit
                                        Sub ExcelSub()
                                            Dim foo As Double
                                        End Sub";
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ImplicitPublicMember"));
        }

        [Test]
        public void MoveFieldCloserToUsageTest()
        {
            const string inputCode = @"Option Explicit 
                    Private bar As String
                    Public Sub Foo()
                        bar = ""test""
                    End Sub";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "MoveFieldCloserToUsage"));
        }

        [Test]
        public void MultipleDeclarationsTest()
        {
            const string inputCode = @"Option Explicit
            Private Sub MySub()
                Dim r As Range, m As Variant
                Set r = Worksheets(""Sheet1"").Range(""A1:C10"")
                m = Application.WorksheetFunction.Min(r)
                MsgBox m
            End Sub ";
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "MultipleDeclarations"));
        }

        [Test]
        public void ObsoleteCallStatementTest()
        {
            const string inputCode = @"Option Explicit
                                        Sub Foo()
                                            Call Foo
                                        End Sub";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", "" + inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ObsoleteCallStatement"));
        }

        [Test]
        public void ObsoleteCommentSyntaxTest()
        {
            const string inputCode = @"Option Explicit 
                    Rem test";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", "" + inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ObsoleteCommentSyntax"));
        }

        [Test]
        public void ObsoleteGlobalTest()
        {
            const string inputCode = @"Option Explicit 
                    Global var1 As Integer";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", "" + inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ObsoleteGlobal"));
        }

        [Test]
        public void ObsoleteLetStatementTest()
        {
            const string inputCode = @"Option Explicit 
                    Public Sub Foo()
                        Dim var1 As Integer
                        Dim var2 As Integer
    
                        Let var2 = var1
                    End Sub                    ";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", "" + inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ObsoleteLetStatement"));
        }

        [Test]
        public void ObsoleteTypeHintTest()
        {
            const string inputCode = @"Option Explicit 
                    Public Foo&";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", "" + inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ObsoleteTypeHint"));
        }

        [Test]
        public void OptionBaseTest()
        {
            const string inputCode = @"Option Explicit 
                    Option Base 1";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", "" + inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "OptionBase"));
        }

        [Test]
        public void OptionExcplicitTest()
        {
            const string inputCode = @"Sub ExcelSub()
                                            Dim foo As Double
                                        End Sub";
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "OptionExplicit"));
        }

        [Test]
        public void ParameterAssignedByValTest()
        {
            const string inputCode = @"Option Explicit
                    Public Sub Foo(ByVal arg1 As String, ByVal arg2 As Integer)
                        arg1 = ""test""
                        arg2 = 9
                    End Sub";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);

            Assert.AreEqual(2, result.VbaCodeIssues.Count);
            Assert.AreEqual(2, result.VbaCodeIssues.Count(x => x.Type == "AssignedByValParameter"));
        }

        [Test]
        public void ParameterCanBeByValTest()
        {
            const string inputCode = @"Option Explicit 
                 Sub Foo(arg1 As Integer)
                End Sub";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ParameterCanBeByVal"));
        }

        [Test]
        public void ParameterNotUsedTest()
        {
            const string inputCode = @"Option Explicit
                  Public Sub Foo(ByVal arg1 As String)
                  End Sub";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count);
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ParameterNotUsed"));
        }

        [Test]
        public void ProcedureCanBeWrittenAsFunctionTest()
        {
            const string inputCode = @"Option Explicit 
                    '@Ignore ParameterCanBeByVal
                    Sub Foo(arg1 As String)
                    End Sub";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", "" + inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ProcedureCanBeWrittenAsFunction"));
        }

        [Test]
        public void ProcedureCanBeWrittenAsFunctiontTest()
        {
            const string inputCode = @"Option Explicit 
                 Sub Foo(arg1 As Integer)
                End Sub";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ProcedureCanBeWrittenAsFunction"));
        }

        [Test]
        public void ProcedureNotUsedTest()
        {
            const string inputCode = @"Option Explicit
            Private Sub MySub()
                Dim r As Range, m As Variant
                Set r = Worksheets(""Sheet1"").Range(""A1:C10"")
                m = Application.WorksheetFunction.Min(r)
                MsgBox m
            End Sub ";
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ProcedureNotUsed"));
        }

        [Test]
        public void SelfAssignedDeclarationTest()
        {
            const string inputCode = @"Option Explicit 
                    Sub Foo()
                        Dim b As New Collection
                    End Sub";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", "" + inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "SelfAssignedDeclaration"));
        }

        [Test]
        public void VariableNotAssigned()
        {
            const string inputCode = @"option explicit
                                        Sub ExcelSub()
                                            Dim foo As Double
                                        End Sub";
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "VariableNotAssigned"));
        }

        [Test]
        public void VariableNotUsed()
        {
            const string inputCode = @"option explicit
                                        Sub ExcelSub()
                                            Dim foo As Double
                                        End Sub";
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "VariableNotUsed"));
        }

        [Test]
        public void VariableNotUsedTest()
        {
            const string inputCode = @"Option Explicit
                    Public Sub Foo(ByVal arg1 As String, ByVal arg2 As Integer)
                     arg1 = ""test""
    
                    Dim var1 As Integer
                    var1 = arg2
                End Sub";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module", inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "VariableNotUsed"));
        }

        [Test]
        public void VariableUnassignedUsageTest()
        {
            const string inputCode = @"Option Explicit
            Private Sub MySub()
                Dim r As Range, m As Variant
                Set r = Worksheets(""Sheet1"").Range(""A1:C10"")
                m = Application.WorksheetFunction.Min(r)
                MsgBox m
            End Sub ";
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);
            Assert.AreEqual(3, result.VbaCodeIssues.Count(x => x.Type == "UnassignedVariableUsage"));
        }

        [Test]
        public void VariableUndeclaredTest()
        {
            const string inputCode = @"Option Explicit
            Private Sub MySub()
                Dim r As Range, m As Variant
                Set r = Worksheets(""Sheet1"").Range(""A1:C10"")
                m = Application.WorksheetFunction.Min(r)
                MsgBox m
            End Sub ";
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", inputCode);
            int Count = result.VbaCodeIssues.Count(x => x.Type == "UndeclaredVariable");

            Assert.AreEqual(3, result.VbaCodeIssues.Count(x => x.Type == "UndeclaredVariable"));
        }

        [Test]
        public void WriteOnlyPropertyTest()
        {
            const string inputCode = @"Option Explicit 
                    Public Property Set Foo(ByVal value As Object)
                    End Property";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", "" + inputCode);

            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "WriteOnlyProperty"));
        }

        [Test, Ignore]
        public void xTester()
        {
            const string inputCode = @"Option Explicit 
                Sub Foo()
                    Dim str As String
                    str = Left$(""test"", 1)
                End Sub
                ";

            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", "" + inputCode);

            if (result.VbaCodeIssues.Count > 0)
            {
                foreach (VbaCodeIssue issue in result.VbaCodeIssues)
                {
                    Debug.WriteLine($"VbaCodeIssues.Types : {issue.Type}");
                }
            }
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "SelfAssignedDeclaration"));
        }
    }
}