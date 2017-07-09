using System.Linq;
using infotron.VbaCodeAnalizer;
using infotron.VbaCodeAnalizer.Inspections;
using NUnit.Framework;

namespace PerfectXL.VbaCodeAnalyzer.UnitTests
{
    [TestFixture]
    public class TypeSpecificTests
    {
        [Test, Ignore]
        public void ApplicationWorksheetFunctionInspectionTest()
        {
            CodeInspection result = CodeAnalizer.Analize(@"
                Option Explicit

                Private Sub MySub()
                    Dim r As Range, m As Variant
                    Set r = Worksheets(""Sheet1"").Range(""A1:C10"")
                    m = Application.WorksheetFunction.Min(r)
                    MsgBox m
                End Sub
                ",
                "Module1",
                "Workbook1.xlsm");
            Assert.AreEqual(1, result.MaintabilityAndReadabilityIssues.Count(x => x.Type == "ApplicationWorksheetFunction"));
        }

        [Test]
        public void ConstantNotUsedInspectionTest()
        {
            CodeInspection result = CodeAnalizer.Analize(@"
                 Option Explicit

                Private Sub MySub()
                    Const c As String = ""foo""
                End Sub
                ",
                "Module1",
                "Workbook1.xlsm");
            Assert.AreEqual(1, result.MaintabilityAndReadabilityIssues.Count(x => x.Type == "ConstantNotUsed"));
        }

        [Test, Ignore]
        public void EmptyStringLiteralInspectionTest()
        {
            CodeInspection result = CodeAnalizer.Analize(@"
                Option Explicit

                Private Sub MySub()
                    Dim s As String
                    s = """"
                    s = s + s
                End Sub
                ",
                "Module1",
                "Workbook1.xlsm");
            Assert.AreEqual(1, result.MaintabilityAndReadabilityIssues.Count(x => x.Type == "EmptyStringLiteral"));
        }

        [Test]
        public void EncapsulatePublicFieldInspectionTest()
        {
            CodeInspection result = CodeAnalizer.Analize(@"
                Option Explicit

                Public x As New Excel.Application

                Private Sub MySub()
                    x.Calculate
                End Sub
                ",
                "Module1",
                "Workbook1.xlsm");
            Assert.AreEqual(1, result.MaintabilityAndReadabilityIssues.Count(x => x.Type == "EncapsulatePublicField"));
        }

        [Test]
        public void FunctionReturnValueNotUsedInspectionTest()
        {
            CodeInspection result = CodeAnalizer.Analize(@"
                Option Explicit

                Private Function MyFunction() As Integer
                    MyFunction = 0
                End Function

                Private Sub MySub()
                    MyFunction
                End Sub
                ",
                "Module1",
                "Workbook1.xlsm");
            Assert.AreEqual(1, result.MaintabilityAndReadabilityIssues.Count(x => x.Type == "FunctionReturnValueNotUsed"));
        }

        // TODO roel Create unit test for every inspection type. Fix ignored unit tests.
    }
}