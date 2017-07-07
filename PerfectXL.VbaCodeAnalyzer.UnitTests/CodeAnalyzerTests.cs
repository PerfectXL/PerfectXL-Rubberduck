using System.Linq;
using infotron.VbaCodeAnalizer;
using infotron.VbaCodeAnalizer.Inspections;
using NUnit.Framework;

namespace PerfectXL.VbaCodeAnalyzer.UnitTests
{
    [TestFixture]
    public class CodeAnalyzerTests
    {
        [Test]
        public void BasicTest()
        {
            CodeInspection result = CodeAnalizer.Analize(@"Option Explicit", "Module1", "Workbook1.xlsm");
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.MaintabilityAndReadabilityIssues.Count);
        }

        [Test]
        public void IssuesTest()
        {
            CodeInspection result = CodeAnalizer.Analize(@"
                Sub MySub()
                    counter = 10
                    For i = 1 To counter
                        MsgBox i
                    Next
                End Sub
                ",
                "Module1",
                "Workbook1.xlsm");
            Assert.AreEqual(7, result.MaintabilityAndReadabilityIssues.Count);
            Assert.AreEqual(1, result.MaintabilityAndReadabilityIssues.Count(x => x.Type == "OptionExplicit"));
        }
    }
}