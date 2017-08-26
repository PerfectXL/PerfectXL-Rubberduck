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

using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.VBEditor.SafeComWrappers;

namespace PerfectXL.VbaCodeAnalyzer.Extensions
{
    public static class VbaCodeModuleExtensions
    {
        private const string PatternExtractVbBaseGuidValues = @"^Attribute \s VB_Base \s = \s "" 0? ( [0-9A-F{}-]+ ) .*";
        private const string PatternTwoGuids = @"(?: \{ [0-9A-F-]+ \} ){2}";

        private static readonly Dictionary<string, ComponentType> GuidToComponentType = new Dictionary<string, ComponentType>
        {
            {"{00020819-0000-0000-C000-000000000046}", ComponentType.Document},
            {"{00020820-0000-0000-C000-000000000046}", ComponentType.Document},
            {"{FCFB3D2A-A0FA-1068-A738-08002B3371B5}", ComponentType.ClassModule}
        };

        public static string StripVbAttributes(this string moduleCode)
        {
            return string.Join("\n",
                moduleCode.Split('\n').SkipWhile(string.IsNullOrWhiteSpace)
                    .SkipWhile(s => Regex.IsMatch(s, @"^ \s* Attribute \s+ VB_\w+ \s+ =", RegexOptions.IgnorePatternWhitespace)));
        }

        public static ComponentType GetModuleType(this string moduleCode)
        {
            string typeGuid = GetTypeGuid(moduleCode);
            if (Regex.IsMatch(typeGuid, PatternTwoGuids, RegexOptions.IgnorePatternWhitespace))
            {
                return ComponentType.UserForm;
            }
            ComponentType value;
            return GuidToComponentType.TryGetValue(typeGuid, out value) ? value : ComponentType.StandardModule;
        }

        private static string GetTypeGuid(string moduleCode)
        {
            return moduleCode.Split('\n')
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => Regex.Replace(s.Trim(), @"\s+", " "))
                .TakeWhile(s => s.StartsWith("Attribute VB_")) /* Take the header part of the module code */
                .Where(s => s.StartsWith(@"Attribute VB_Base = """))
                .Select(ExtractVbBaseGuidValues)
                .FirstOrDefault() ?? "";
        }

        private static string ExtractVbBaseGuidValues(string s)
        {
            return Regex.Replace(s, PatternExtractVbBaseGuidValues, "$1", RegexOptions.IgnorePatternWhitespace);
        }
    }
}