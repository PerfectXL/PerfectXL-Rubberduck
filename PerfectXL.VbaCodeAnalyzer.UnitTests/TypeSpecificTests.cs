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
            CodeInspectionResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1",
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
        public void ConstantNotUsedInspectionTest()
        {
            CodeInspectionResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1",
                @"
                 Option Explicit

                Private Sub MySub()
                    Const c As String = ""foo""
                End Sub
                ");
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "ConstantNotUsed"));
        }

        [Test, Ignore]
        public void EmptyStringLiteralInspectionTest()
        {
            CodeInspectionResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1",
                @"
                Option Explicit

                Private Sub MySub()
                    Dim s As String
                    s = """"
                    s = s + s
                End Sub
                ");
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "EmptyStringLiteral"));
        }

        [Test]
        public void EncapsulatePublicFieldInspectionTest()
        {
            CodeInspectionResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1",
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
        public void FunctionReturnValueNotUsedInspectionTest()
        {
            CodeInspectionResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1",
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

        // TODO roel Create unit test for every inspection type. Fix ignored unit tests.
    }
}
