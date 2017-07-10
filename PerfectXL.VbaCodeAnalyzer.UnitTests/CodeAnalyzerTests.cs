using System.Linq;
using NUnit.Framework;
using PerfectXL.VbaCodeAnalyzer.Inspection;

namespace PerfectXL.VbaCodeAnalyzer.UnitTests
{
    [TestFixture]
    public class CodeAnalyzerTests
    {
        [Test]
        public void BasicTest()
        {
            CodeInspectionResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", @"Option Explicit");
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.VbaCodeIssues.Count);
        }

        [Test]
        public void IssuesTest()
        {
            CodeInspectionResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1",
                @"
                Sub MySub()
                    counter = 10
                    For i = 1 To counter
                        MsgBox i
                    Next
                End Sub
                ");
            Assert.AreEqual(7, result.VbaCodeIssues.Count);
            Assert.AreEqual(1, result.VbaCodeIssues.Count(x => x.Type == "OptionExplicit"));
        }
    }
}