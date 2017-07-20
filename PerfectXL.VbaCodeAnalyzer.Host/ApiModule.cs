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
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using Nancy;
using Nancy.ModelBinding;
using NLog;
using PerfectXL.VbaCodeAnalyzer.Host.Models;
using PerfectXL.VbaCodeAnalyzer.Inspection;

namespace PerfectXL.VbaCodeAnalyzer.Host
{
    public class ApiModule : NancyModule
    {
        private static readonly Logger MyLogger = LogManager.GetCurrentClassLogger();

        public ApiModule()
        {
            Get["/"] = parameters => $@"<html><head><title>{Program.Name}</title></head>
                <body><h1 style=""font:700 18px sans-serif;text-align:center;"">{Program.Name}</h1>
                <p style=""text-align:center;""><img src=""data:image/png;base64,{GetResource()}"" /></p></body></html>";

            Post["/v1/analyze/project"] = x => AnalyzeProject(x);
        }

        private IList<CodeInspectionResult> AnalyzeProject(dynamic parameters)
        {
            VbaProject model;
            try
            {
                model = this.Bind<VbaProject>(new BindingConfig {BodyOnly = true});
            }
            catch (Exception exception)
            {
                MyLogger.Error(exception);
                throw;
            }
            IList<CodeInspectionResult> results = AnalyzeProject(model);
            return results;
        }

        private static IList<CodeInspectionResult> AnalyzeProject(VbaProject model)
        {
            MyLogger.Debug($"Analyzing {model.FileName} with {model.VbaModules.Count} VBA modules.");
            Stopwatch sw = Stopwatch.StartNew();
            IList<CodeInspectionResult> results;
            try
            {
                results = new CodeAnalyzer(model.FileName).Run(model.VbaModules.ToDictionary(x => x.Name, x => x.Code));
            }
            catch (Exception exception)
            {
                MyLogger.Error(exception);
                throw;
            }
            MyLogger.Debug($"Analysis took {sw.ElapsedMilliseconds}ms.");
            return results;
        }

        private static string GetResource()
        {
            using (var ms = new MemoryStream())
            {
                Assembly.GetExecutingAssembly().GetManifestResourceStream("PerfectXL.VbaCodeAnalyzer.Host.Rubberduck.png")?.CopyTo(ms);
                return Convert.ToBase64String(ms.ToArray());
            }
        }
    }
}