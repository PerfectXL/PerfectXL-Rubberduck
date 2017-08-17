using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Antlr4.Runtime.Misc;
using PerfectXL.VbaCodeAnalyzer.Inspection;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace PerfectXL.VbaCodeAnalyzer.Macro
{
    public static class MacroInspector
    {
        public static List<VbaCodeIssue> Run(RubberduckParserState vbaObject)
        {
            return QualifyMacro(vbaObject);
        }

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

        private static List<VbaCodeIssue> QualifyMacro(RubberduckParserState vbaObject)
        {
            var macroIssues = new List<VbaCodeIssue>();

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
                var hasSelectFollowedByActive = Regex.Matches(macroText, @"\bSelect\s*\n\s*Active\w+").Count > 0;

                var isRecorded2 = (hasMacroComment || hasSelectFollowedByActive) && !hasDimStatements;

                if (isRecorded2)
                {
                    IInspectionResult item = null;
                    var macroIssue = new VbaCodeIssue(item , "filename", "modulename")
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
                    IInspectionResult item = null;
                    var macroIssue = new VbaCodeIssue(item, "filename", "modulename")
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
    }
}
