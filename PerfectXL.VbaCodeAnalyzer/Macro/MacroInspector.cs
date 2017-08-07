using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace PerfectXL.VbaCodeAnalyzer.Macro
{
    public static class MacroInspector
    {
        private static readonly List<MacroState> MacroStateInspection = new List<MacroState>();

        public static List<MacroState> Run(RubberduckParserState vbaObject)
        {
            var macroTermsList = Inspect(vbaObject);
            var macroIsPredefined = false;
            var macroName = "";

            var terms = MacroTerm.Rates();

            var functionList = macroTermsList.GroupBy(grp => grp.Function).Select(g => g.Key).Distinct().ToList();

            foreach (var function in functionList)
            {
                var macros = macroTermsList.Where(l => l.Function == function).ToList();

                foreach (var macro in macros)
                {
                    if (macro.Term.Length > 2 && macro.Term.StartsWith("xl"))
                    {
                        MacroTypeToCache(macroName, true);
                        break;
                    }
                    var termRate = terms.Select(x => x).FirstOrDefault(x => x.Term == macro.Term);

                    if (termRate == null) continue;

                    macroName = macro.Function;

                    macroIsPredefined = (macro.Percentage > termRate.Rate);
                }
                MacroTypeToCache(macroName, macroIsPredefined);
            }
            return MacroStateInspection;
        }

        private static IEnumerable<MacroTermPresenter> Inspect(RubberduckParserState vbaObject)
        {
            return MacroTermsCounter(CreateDeclarationList(vbaObject));
        }

        private static IEnumerable<Declaration> CreateDeclarationList(RubberduckParserState vbaObject)
        {
            var declarations = new List<Declaration>();

            var allDeclarations = vbaObject.AllDeclarations
                .Where(l => l.DeclarationType != DeclarationType.Project && l.DeclarationType != DeclarationType.ProceduralModule && l.DeclarationType != DeclarationType.Procedure)
                .GroupBy(grp => grp.ParentDeclaration)
                .SelectMany(g => g.OrderBy(grp => grp.Selection.StartLine)).ToArray();

            declarations.AddRange(allDeclarations);

            var allUnresolvedMemberDeclarations = vbaObject.AllUnresolvedMemberDeclarations
                .Where(l => l.DeclarationType == DeclarationType.UnresolvedMember).ToArray();

            declarations.AddRange(allUnresolvedMemberDeclarations);

            return declarations;
        }

        private static IEnumerable<MacroTermPresenter> MacroTermsCounter(IEnumerable<Declaration> declarations)
        {
            var presenterList = new List<MacroTermPresenter>();

            CounterFunctionTerm.Clear();

            var functionList = declarations.GroupBy(grp => grp.ParentDeclaration.IdentifierName).Select(g => g.Key).Distinct().ToList();

            foreach (var function in functionList)
            {
                var functionTerms = declarations.Where(l => l.ParentDeclaration.IdentifierName == function).Select(x => x).OrderBy(x => x.IdentifierName).ToList();

                presenterList.AddRange(CreatePresenterList(functionTerms));
            }
            return presenterList;
        }

        private static IEnumerable<MacroTermPresenter> CreatePresenterList(IEnumerable<Declaration> declarations)
        {
            var totalTermCounter = 0;

            var presenterList = new List<MacroTermPresenter>();

            foreach (var declaration in declarations)
            {
                var termCounter = 1;

                if (declaration.IsUndeclared)
                {
                    termCounter = declaration.References.GroupBy(grp => grp.Selection).Select(g => g.Key).Distinct().Count();
                }

                totalTermCounter = totalTermCounter + termCounter;

                FunctionTermCounter(declaration.ParentDeclaration.IdentifierName, termCounter);

                var termPresenter = new MacroTermPresenter
                {
                    Module = declaration.ComponentName,
                    Function = declaration.ParentDeclaration.IdentifierName,
                    Term = declaration.IdentifierName,
                    Repeat = termCounter,
                    Listed = IsListedTerm(declaration.IdentifierName)
                };

                var presenterInList = presenterList.Find(x => x.Term == termPresenter.Term && x.Function == termPresenter.Function);

                if (presenterInList == null)
                {
                    presenterList.Add(termPresenter);
                }
                else
                {
                    presenterInList.Repeat = presenterInList.Repeat + termPresenter.Repeat;
                }
            }

            CalculatePercentage(presenterList);

            return presenterList;
        }

        private static readonly Dictionary<string, int> CounterFunctionTerm = new Dictionary<string, int>();
        private static void FunctionTermCounter(string macro, int termcounter)
        {
            if (CounterFunctionTerm.ContainsKey(macro))
            {
                var value = CounterFunctionTerm[macro];
                CounterFunctionTerm[macro] = value + termcounter;
            }
            else
            {
                CounterFunctionTerm.Add(macro, termcounter);
            }
        }

        private static bool IsListedTerm(string term)
        {
            return (MacroTerm.List().Where(s => s == term).ToList().Count != 0);
        }

        private static void CalculatePercentage(IEnumerable<MacroTermPresenter> presenterList)
        {
            foreach (var item in presenterList)
            {
                var count = CounterFunctionTerm[item.Function];
                item.Percentage = Math.Round((decimal)item.Repeat / count, 3);
            }
        }

        private static void MacroTypeToCache(string function, bool type)
        {
            var macroStateItem = MacroStateInspection.Find(x => x.Name == function);

            var macrotype = ((type) ? "Predefined" : "Recorded");
            if (macroStateItem == null)
            {
                MacroStateInspection.Add(new MacroState { Name = function, State = macrotype });
            }
            else
            {
                macroStateItem.State = macrotype;
            }
        }
    }
}
