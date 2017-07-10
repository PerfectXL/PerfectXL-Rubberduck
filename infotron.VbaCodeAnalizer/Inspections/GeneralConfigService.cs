using System;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Settings;

namespace infotron.VbaCodeAnalizer.Inspections {
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

        public event EventHandler<ConfigurationChangedEventArgs> SettingsChanged;

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