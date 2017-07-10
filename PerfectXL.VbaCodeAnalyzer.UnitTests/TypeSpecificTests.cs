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