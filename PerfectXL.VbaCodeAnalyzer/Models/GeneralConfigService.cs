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
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Settings;

namespace PerfectXL.VbaCodeAnalyzer.Models
{
    internal class GeneralConfigService : IGeneralConfigService
    {
        private readonly Configuration _configuration;

        public GeneralConfigService(IInspectionModel inspection)
        {
            _configuration = CreateConfiguration(inspection);
        }

        public Configuration LoadConfiguration()
        {
            return _configuration;
        }

        public void SaveConfiguration(Configuration toSerialize)
        {
            throw new NotImplementedException();
        }

#pragma warning disable 67
        public event EventHandler<ConfigurationChangedEventArgs> SettingsChanged;
#pragma warning restore 67

        public Configuration GetDefaultConfiguration()
        {
            throw new NotImplementedException();
        }

        private static Configuration CreateConfiguration(IInspectionModel inspection)
        {
            var settings = new CodeInspectionSettings();
            settings.CodeInspections.Add(new CodeInspectionSetting {Description = inspection.Description, Severity = inspection.Severity});
            return new Configuration {UserSettings = new UserSettings(null, null, null, settings, null, null, null)};
        }
    }
}
