﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using NUnit.Framework;

namespace PerfectXL.VbaCodeAnalyzer.UnitTests
{
    [TestFixture]
    public class ParserTest
    {
        [Test]
        public void TestParser()
        {
            var codeUrenregistratie = CodeExtractor(@"Macros\" + TestFileNames.UserDefined_Macro_1 + ".txt");

            if (codeUrenregistratie != string.Empty)
            {
                var macroTypeCache = new CodeAnalyzer("Workbook1.xlsm").RankMacro("Module1", codeUrenregistratie);

                foreach (var macrotype in macroTypeCache)
                {
                    var mcrotype = "Macro: " + macrotype.Name + " is a " + macrotype.State + " macro";
                    Debug.WriteLine(mcrotype);
                }
            }
        }

        [Test]
        public void TestFilter()
        {
            var file = @"C:\Users\HarveyBouva\Projects\PerfectXL\PerfectXL-Rubberduck\PerfectXL-Rubberduck\MacroTermRating.xml";

            XDocument xdoc = null;

            using (XmlReader xr = XmlReader.Create(file))
            {
                xdoc = XDocument.Load(xr);
            }


            // XmlTextReader reader = new XmlTextReader(file);

            //var xml = XDocument.Load(file);
            XElement xelement = XElement.Load(file);

            // IEnumerable<XElement> employees = xelement.Elements();


            //foreach (var employee in employees)
            //{
            //    Debug.WriteLine(employee.Element("Name").Value);
            //}


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
