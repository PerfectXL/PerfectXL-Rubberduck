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
    public class CodeAnalyzerTests
    {
        [Test]
        public void BasicTest()
        {
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", @"Option Explicit");
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.VbaCodeIssues.Count);
        }

        [Test]
        public void IssuesTest()
        {
            CodeAnalyzerResult result = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1",
                @"Sub MySub()
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