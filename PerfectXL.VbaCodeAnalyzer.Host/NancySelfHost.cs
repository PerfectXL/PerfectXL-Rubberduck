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
using System.Configuration;
using Nancy.Hosting.Self;

namespace PerfectXL.VbaCodeAnalyzer.Host
{
    internal class NancySelfHost
    {
        private static readonly int Port = int.Parse(ConfigurationManager.AppSettings["port"]);
        private readonly Uri _uri = new UriBuilder {Scheme = "http", Port = Port, Host = "localhost"}.Uri;
        private NancyHost _nancyHost;

        public void Start()
        {
            _nancyHost = new NancyHost(_uri);
            Console.WriteLine($"Running {Program.Name} on {_uri}.");
            _nancyHost.Start();
        }

        public void Stop()
        {
            _nancyHost.Stop();
            Console.WriteLine($"Stopped {Program.Name}.");
        }
    }
}
