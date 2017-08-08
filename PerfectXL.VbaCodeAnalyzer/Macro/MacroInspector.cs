using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Antlr4.Runtime.Misc;
using PerfectXL.VbaCodeAnalyzer.Inspection;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace PerfectXL.VbaCodeAnalyzer.Macro
{
    public static class MacroInspector
    {
        private static readonly List<VbaMacroIssue> MacroStateInspection = new List<VbaMacroIssue>();

        public static List<VbaMacroIssue> Run(RubberduckParserState vbaObject)
        {
            return QualifyMacro(vbaObject);
        }

        //private static IEnumerable<MacroTermPresenter> Inspect(RubberduckParserState vbaObject)
        //{
        //   var macroIssues = QualifyMacro(vbaObject);

        //    return MacroTermsCounter(CreateDeclarationList(vbaObject));
        //}

        private static IEnumerable<Declaration> CreateDeclarationList(RubberduckParserState vbaObject)
        {
            var declarations = new List<Declaration>();

            var allDeclarations = vbaObject.AllDeclarations
                .Where(l => l.DeclarationType != DeclarationType.Project && l.DeclarationType != DeclarationType.ProceduralModule)
                .GroupBy(grp => grp.ParentDeclaration)
                .SelectMany(g => g.OrderBy(grp => grp.Selection.StartLine)).ToArray();

            declarations.AddRange(allDeclarations);

            var allUnresolvedMemberDeclarations = vbaObject.AllUnresolvedMemberDeclarations
                .Where(l => l.DeclarationType == DeclarationType.UnresolvedMember).ToArray();

            declarations.AddRange(allUnresolvedMemberDeclarations);

            return declarations;
        }

        private static List<VbaMacroIssue> QualifyMacro(RubberduckParserState vbaObject)
        {
            // Het is opgenomen macro als
            // Er nooit het woord “DIM spatie” gebruikt wordt
            // Er in het eerste commentaar het woord “spatie Macro” staat
            // Als er minimaal 1 maal in de code staat “Select white space Active”
            // returns a list with recorded macro's

            var macroIssues = new List<VbaMacroIssue>();

            var hasMacroComment = false;

            var comments = vbaObject.AllComments;

            foreach (var comment in comments)
            {
                if (comment.CommentText != string.Empty)
                {
                    hasMacroComment = Regex.Matches(comment.CommentText, @"\s*\w+ Macro").Count > 0;
                }
            }

            var declarations = CreateDeclarationList(vbaObject);

            var functionDeclarations = declarations.Where(l => l.DeclarationType == DeclarationType.Function || l.DeclarationType == DeclarationType.Procedure).ToArray();

            foreach (var function in functionDeclarations)
            {
                var context = function.Context.Start.InputStream;
                var startIndex = function.Context.Start.StartIndex;
                var endIndex = function.Context.Stop.StopIndex;
                var macroText = context.GetText(new Interval(startIndex, endIndex));

                var hasDimStatements = Regex.Matches(macroText, @"(^|\n)\s*Dim").Count > 0;
                var hasSelectFollowedByActive = Regex.Matches(macroText, @"Select\s*\n\s*Active").Count > 0;

                var isRecorded = (hasMacroComment || hasSelectFollowedByActive || !hasDimStatements);

                if (isRecorded)
                {
                    var macroIssue = new VbaMacroIssue
                    {
                        Type = "RecordedMacro",
                        Name = function.IdentifierName,
                        Description = $"The {function.DeclarationType} '{function.IdentifierName}’ looks like it was recorded.",
                        Meta = $"Recorder macros are often badly written and highly context - dependent.Consider rewriting this {function.DeclarationType}.",
                        Line = function.Selection.StartLine,
                        Column = function.Selection.StartColumn,
                        ModuleName = function.ComponentName,
                        Severity = function.Accessibility.ToString()
                    };
                    macroIssues.Add(macroIssue);
                }

                var hasWorkbookOpen = Regex.Matches(macroText, @"\s*\w+ Workbook_Open").Count > 0;
                var hasWorkbookOpenen = Regex.Matches(macroText, @"\s*\w+ Workbook_Openen").Count > 0;
                var hasWorkbookBeforeSave = Regex.Matches(macroText, @"\s*\w+ Workbook_BeforeSave").Count > 0;
                var hasWorkSheetOpen = Regex.Matches(macroText, @"\s*\w+ Auto_Open").Count > 0;
                var hashWorkbookBeforeClose = Regex.Matches(macroText, @"\s*\w+ Workbook_BeforeClose").Count > 0;
                var hashWorksheetSelectionChange = Regex.Matches(macroText, @"\s*\w+ Worksheet_SelectionChange").Count > 0;
                var hasWorksheetActivate = Regex.Matches(macroText, @"\s*\w+ Worksheet_Activate").Count > 0;
                var hasWorksheetChange = Regex.Matches(macroText, @"\s*\w+ Worksheet_Change").Count > 0;
                var hasAutoOpen = Regex.Matches(macroText, @"\s*\w+ Auto_Open").Count > 0; 
                var hasRun = Regex.Matches(macroText, @"\s*\w+ Run").Count > 0;

                var isAuto = hasWorkbookOpen || hasWorkbookOpenen || 
                    hasWorkbookBeforeSave || hasWorkSheetOpen || hashWorkbookBeforeClose || 
                    hashWorksheetSelectionChange || hasWorksheetActivate || hasWorksheetChange || 
                    hasAutoOpen || hasRun;

                if (isAuto)
                {
                    var macroIssue = new VbaMacroIssue
                    {
                        Type = "AutoOpenMacro",
                        Name = function.IdentifierName,
                        Description = $"The {function.DeclarationType} {function.IdentifierName} content code to run Automatically.",
                        Meta = $"Auto-Run macros are often dangerous PerfectXl recomand to review the {function.DeclarationType} {function.IdentifierName}.",
                        Line = function.Selection.StartLine,
                        Column = function.Selection.StartColumn,
                        ModuleName = function.ComponentName,
                        Severity = function.Accessibility.ToString()
                    };
                    macroIssues.Add(macroIssue);
                }
            }
            return macroIssues;
        }

        //private static IEnumerable<MacroTermPresenter> MacroTermsCounter(IEnumerable<Declaration> declarations)
        //{
        //    var presenterList = new List<MacroTermPresenter>();

        //    CounterFunctionTerm.Clear();

        //    var functionList = declarations.GroupBy(grp => grp.ParentDeclaration.IdentifierName).Select(g => g.Key).Distinct().ToList();

        //    foreach (var function in functionList)
        //    {
        //        var functionTerms = declarations.Where(l => l.ParentDeclaration.IdentifierName == function).Select(x => x).OrderBy(x => x.IdentifierName).ToList();

        //        presenterList.AddRange(CreatePresenterList(functionTerms));
        //    }
        //    return presenterList;
        //}

        //private static IEnumerable<MacroTermPresenter> CreatePresenterList(IEnumerable<Declaration> declarations)
        //{
        //    var totalTermCounter = 0;

        //    var presenterList = new List<MacroTermPresenter>();

        //    foreach (var declaration in declarations)
        //    {
        //        var termCounter = 1;

        //        if (declaration.IsUndeclared)
        //        {
        //            termCounter = declaration.References.GroupBy(grp => grp.Selection).Select(g => g.Key).Distinct().Count();
        //        }

        //        totalTermCounter = totalTermCounter + termCounter;

        //        FunctionTermCounter(declaration.ParentDeclaration.IdentifierName, termCounter);

        //        var termPresenter = new MacroTermPresenter
        //        {
        //            Module = declaration.ComponentName,
        //            Function = declaration.ParentDeclaration.IdentifierName,
        //            Term = declaration.IdentifierName,
        //            Repeat = termCounter,
        //            Listed = IsListedTerm(declaration.IdentifierName)
        //        };

        //        var presenterInList = presenterList.Find(x => x.Term == termPresenter.Term && x.Function == termPresenter.Function);

        //        if (presenterInList == null)
        //        {
        //            presenterList.Add(termPresenter);
        //        }
        //        else
        //        {
        //            presenterInList.Repeat = presenterInList.Repeat + termPresenter.Repeat;
        //        }
        //    }

        //    CalculatePercentage(presenterList);

        //    return presenterList;
        //}

        //private static readonly Dictionary<string, int> CounterFunctionTerm = new Dictionary<string, int>();
        //private static void FunctionTermCounter(string macro, int termcounter)
        //{
        //    if (CounterFunctionTerm.ContainsKey(macro))
        //    {
        //        var value = CounterFunctionTerm[macro];
        //        CounterFunctionTerm[macro] = value + termcounter;
        //    }
        //    else
        //    {
        //        CounterFunctionTerm.Add(macro, termcounter);
        //    }
        //}

        //private static bool IsListedTerm(string term)
        //{
        //    return (MacroTerm.List().Where(s => s == term).ToList().Count != 0);
        //}

        //private static void CalculatePercentage(IEnumerable<MacroTermPresenter> presenterList)
        //{
        //    foreach (var item in presenterList)
        //    {
        //        var count = CounterFunctionTerm[item.Function];
        //        item.Percentage = Math.Round((decimal)item.Repeat / count, 3);
        //    }
        //}

        //private static void MacroTypeToCache(string function, bool type)
        //{
        //    var macroStateItem = MacroStateInspection.Find(x => x.Name == function);

        //    var macrotype = ((type) ? "Predefined" : "Recorded");
        //    if (macroStateItem == null)
        //    {
        //        MacroStateInspection.Add(new VbaMacroIssue { Name = function, State = macrotype });
        //    }
        //    else
        //    {
        //        macroStateItem.State = macrotype;
        //    }
        //}

    }
}
