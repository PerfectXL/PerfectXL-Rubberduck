﻿using System.IO;
using NUnit.Framework;

namespace PerfectXL.VbaCodeAnalyzer.UnitTests
{
    [TestFixture]
    public class ParserTest
    {
        [Test]
        public void TestParser()
        {
            var codeUrenregistratie = CodeExtractor(@"Macros\" + TestFileNames.Predefined_Planning_Urenregistratie_balken + ".txt");

            if (codeUrenregistratie == string.Empty) return;
            var macroTypeCache = new CodeAnalyzer("Workbook1.xlsm").RankMacro("Module1", codeUrenregistratie);

            Assert.AreEqual(1, macroTypeCache.Count);
        }

        [Test]
        public void AnalyzeTest()
        {
            var codeUrenregistratie = CodeExtractor(@"Macros\" + TestFileNames.UserDefined_Macro_1 + ".txt");

            if (codeUrenregistratie == string.Empty) return;
            var analyzeResult = new CodeAnalyzer("Workbook1.xlsm").AnalyzeModule("Module1", codeUrenregistratie);

            Assert.AreEqual(13, analyzeResult.VbaCodeIssues.Count);
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
        
        public enum TestFileNames
        {
            adreslijst_sorteren,
            DCF_WONEN_Blad9,
            DCF_WONEN_foto,
            DCF_WONEN_getallen_nederlands,
            DCF_WONEN_getallen_uitschrijven,
            Predefined_Casheflow_BilledSalesInlezenBilledSales,
            Predefined_Casheflow_Module1,
            Predefined_Casheflow_SubDTVernieuwenDBForecast,
            Predefined_Casheflow_SubDTVernieuwenDBForecastvsActuals,
            Predefined_Casheflow_SubDTVernieuwenHoofdblad,
            Predefined_Hurdles_Module1,
            Predefined_Hurdles_Module2,
            Predefined_Hurdles_Sheet1,
            Predefined_Hurdles_Sheet2,
            Predefined_Planning_Instelling,
            Predefined_Planning_Openen,
            Predefined_Planning_Roosterplanning,
            Predefined_Planning_Roosterplanning_bereken_totalen,
            Predefined_Planning_Roosterplanning_laden_beschikbaarheidsplanning,
            Predefined_Planning_Roosterplanning_uren_plannen,
            Predefined_Planning_Urenregistratie,
            Predefined_Planning_Urenregistratie_balken,
            UserDefined_Macro_1,
            UserDefined_Macro_1_Eerste_Opname,
            UserDefined_Macro_2
        }

    }
}
