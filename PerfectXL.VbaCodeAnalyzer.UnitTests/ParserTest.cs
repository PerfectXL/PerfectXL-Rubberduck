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
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using NUnit.Framework;
using PerfectXL.VbaCodeAnalyzer.Extensions;
using PerfectXL.VbaCodeAnalyzer.Inspection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace PerfectXL.VbaCodeAnalyzer.UnitTests
{
    // TODO These are not unit tests. Remove or rewrite.
    // TODO Class name does not reflect purpose of this module.
    [TestFixture]
    public class ParserTest
    {
        [Test]
        public void TestParser()
        {
            string codeUrenregistratie = CodeExtractor(@"Planning_uren_v1.4_20160824\Modules\Urenregistratie.txt");
            string codeRoosterplanning = CodeExtractor(@"Planning_uren_v1.4_20160824\Modules\Roosterplanning.txt");

            IEnumerable<MacroTermPresenter> urenregistratieTermList = null;
            IEnumerable<MacroTermPresenter> roosterplanningTermList = null;

            if (codeUrenregistratie != string.Empty)
            {
                RubberduckParserState vbaObject = new CodeAnalyzer("Workbook1.xlsm").Parse("Module1", codeUrenregistratie).ParserState;
                urenregistratieTermList = Analize(vbaObject);
            }

            if (codeRoosterplanning != string.Empty)
            {
                RubberduckParserState vbaObject = new CodeAnalyzer("Workbook1.xlsm").Parse("Module1", codeRoosterplanning).ParserState;
                roosterplanningTermList = Analize(vbaObject);
            }
        }

        public static IEnumerable<MacroTermPresenter> Analize(RubberduckParserState vbaObject)
        {
            var unresolvedMemberDeclarations = vbaObject.AllDeclarations.GroupBy(grp => grp.ParentDeclaration).SelectMany(g => g.OrderBy(grp => grp.Selection.StartLine)).ToArray();

            return MacroTermsCounter(unresolvedMemberDeclarations, MacroTerms.List());
        }

        private static IEnumerable<MacroTermPresenter> MacroTermsCounter(IEnumerable<Declaration> declarations, IEnumerable<string> terms)
        {
            var presenterList = new List<MacroTermPresenter>();
            foreach (var term in terms)
            {
                var declarationQuery = declarations.Where(x => x.IdentifierName == term).Select(item => new { item }).ToList();

                foreach (var declaration in declarationQuery)
                {
                    var termPresenter = new MacroTermPresenter
                    {
                        Module = declaration.item.ComponentName,
                        Function = declaration.item.ParentDeclaration.IdentifierName,
                        Term = term,
                        Repeat = declaration.item.References.Count()
                    };
                    presenterList.Add(termPresenter);
                }
            }

            CalculatePercentage(presenterList);

            return presenterList;
        }

        private static void CalculatePercentage(IEnumerable<MacroTermPresenter> presenters)
        {
            var macroTermPresenters = presenters as IList<MacroTermPresenter> ?? presenters.ToList();
            var sum = (double)macroTermPresenters.Select(r => r.Repeat).Sum();

            foreach (var present in macroTermPresenters)
            {
                present.Percentage = Math.Round(100 * present.Repeat / sum);
            }
        }

        private static string CodeExtractor(string path)
        {
            var vbaCode = "";
            const string filepath = @"C:\Users\HarveyBouva\Projects\PerfectXL\SampleFiles\";

            if (!File.Exists(filepath + path)) return vbaCode;
            using (var sr = new StreamReader(filepath + path))
            {
                vbaCode = sr.ReadToEnd();
            }
            return vbaCode;
        }

        public static int CountStringOccurrences(string text, string pattern)
        {
            var count = 0;
            var i = 0;
            while ((i = text.IndexOf(pattern, i, StringComparison.Ordinal)) != -1)
            {
                i += pattern.Length;
                count++;
            }
            return count;
        }

        public class MacroTermPresenter
        {
            public string Module { get; set; }
            public string Function { get; set; }
            public string Term { get; set; }
            public int Repeat { get; set; }
            public double Percentage { get; set; }
        }

        public static class MacroTerms
        {
            private static readonly List<string> _terms = new List<string>();

            static MacroTerms()
            {
                _terms.AddRange(new List<string>
                {
                    "Activate",
                    "ActiveChart",
                    "ActiveSheet",
                    "ActiveWorkbook",
                    "Add",
                    "AllowMultiSelect",
                    "Application",
                    "Apply",
                    "AutoClose",
                    "AutoExec",
                    "AutoExit",
                    "AutoFill",
                    "AutoNew",
                    "AutoOpen",
                    "Clear",
                    "Close",
                    "Copy",
                    "CutCopyMode",
                    "Delete",
                    "Display3DShading",
                    "DisplayFullScreen",
                    "DisplayHeadings",
                    "DropDownLines",
                    "Header",
                    "Insert",
                    "LinkedCell",
                    "ListFillRange",
                    "MatchVase",
                    "MsgBox",
                    "msoFileOpen",
                    "Open",
                    "Orientation",
                    "PastSpecial",
                    "Protect",
                    "Range",
                    "Save",
                    "ScreenUpdating",
                    "Select",
                    "Selection",
                    "SelectionChange",
                    "SetRange",
                    "Sheets",
                    "Show",
                    "SortMethod",
                    "Unprotect",
                    "Values",
                    "Windows",
                    "Workbook_Open",
                    "XValues",
                    "Zoom"
                });
            }

            public static List<string> List()
            {
                return _terms;
            }
        }

    }
}
