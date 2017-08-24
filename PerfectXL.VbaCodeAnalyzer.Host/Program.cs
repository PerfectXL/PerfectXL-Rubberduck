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
using System.Reflection;
using Nancy.Json;
using Topshelf;

namespace PerfectXL.VbaCodeAnalyzer.Host
{
    internal class Program
    {
        public static string Name { get; } = "PerfectXL.VbaCodeAnalyzer.Host";
        public static Version Version { get; } = Assembly.GetExecutingAssembly().GetName().Version;

        private static void Main()
        {
            JsonSettings.MaxJsonLength = int.MaxValue;

            HostFactory.Run(x =>
            {
                x.Service<NancySelfHost>(s =>
                {
                    s.ConstructUsing(name => new NancySelfHost());
                    s.WhenStarted(tc => tc.Start());
                    s.WhenStopped(tc => tc.Stop());
                });

                x.RunAsLocalSystem();
                x.SetDescription(Name);
                x.SetDisplayName(Name);
                x.SetServiceName(Name);
            });
        }
    }
}