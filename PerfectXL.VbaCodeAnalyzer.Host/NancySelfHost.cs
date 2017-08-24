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
using System.Configuration;
using System.Linq;
using Nancy;
using Nancy.Hosting.Self;
using Nancy.TinyIoc;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using NLog;
using PerfectXL.VbaCodeAnalyzer.Inspection;
using PerfectXL.VbaCodeAnalyzer.Parsing;

namespace PerfectXL.VbaCodeAnalyzer.Host
{
    internal class NancySelfHost
    {
        private static readonly Logger MyLogger = LogManager.GetCurrentClassLogger();
        private static readonly int Port = int.Parse(ConfigurationManager.AppSettings["port"]);
        private readonly Uri _uri = new UriBuilder {Scheme = "http", Port = Port, Host = "localhost"}.Uri;
        private NancyHost _nancyHost;

        public void Start()
        {
            _nancyHost = new NancyHost(_uri, new Bootstrapper());
            MyLogger.Info($"Running {Program.Name} on {_uri}.");
            _nancyHost.Start();
        }

        public void Stop()
        {
            _nancyHost.Stop();
            MyLogger.Info($"Stopped {Program.Name}.");
        }

        #region Nancy host configuration
        private class Bootstrapper : DefaultNancyBootstrapper
        {
            protected override void ConfigureApplicationContainer(TinyIoCContainer container)
            {
                base.ConfigureApplicationContainer(container);
                container.Register<JsonSerializer, CustomJsonSerializer>();
            }
        }

        // ReSharper disable once ClassNeverInstantiated.Local
        private class CustomJsonSerializer : JsonSerializer
        {
            public CustomJsonSerializer()
            {
                Initialize();
            }

            private void Initialize()
            {
                PreserveReferencesHandling = PreserveReferencesHandling.All;
                TypeNameHandling = TypeNameHandling.Auto;
                TypeNameAssemblyFormatHandling = TypeNameAssemblyFormatHandling.Simple;
                SerializationBinder = new CustomSerializationBinder();
            }
        }

        private class CustomSerializationBinder : ISerializationBinder
        {
            private static readonly List<Type> AllowedTypes = new List<Type>
            {
                typeof(CodeAnalyzerResult),
                typeof(VbaCodeIssue),
                typeof(VbaParseTree),
                typeof(ErrorNode),
                typeof(Interval),
                typeof(Rule),
                typeof(Token),
                typeof(UnkownNode)
            };

            public Type BindToType(string assemblyName, string typeName)
            {
                return AllowedTypes.FirstOrDefault(x => x.Name == typeName);
            }

            public void BindToName(Type serializedType, out string assemblyName, out string typeName)
            {
                assemblyName = "Vba";
                typeName = serializedType.Name;
            }
        }
        #endregion Nancy host configuration
    }
}