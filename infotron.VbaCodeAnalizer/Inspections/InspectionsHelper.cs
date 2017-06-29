using System.Linq;
using Moq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Settings;

namespace infotron.VbaCodeAnalizer.Inspections
{
   public class InspectionsHelper
    {
        public static IInspector GetInspector(IInspection inspection, params IInspection[] otherInspections)
        {
            return new Rubberduck.Inspections.Rubberduck.Inspections.Inspector(GetSettings(inspection), otherInspections.Union(new[] { inspection }));
        }

        public static IGeneralConfigService GetSettings(IInspection inspection)
        {
            var settings = new Mock<IGeneralConfigService>();
            var config = GetTestConfig(inspection);
            settings.Setup(x => x.LoadConfiguration()).Returns(config);

            return settings.Object;
        }

        private static Configuration GetTestConfig(IInspection inspection)
        {
            var settings = new CodeInspectionSettings();
            settings.CodeInspections.Add(new CodeInspectionSetting
            {
                Description = inspection.Description,
                Severity = inspection.Severity
            });
            return new Configuration
            {
                UserSettings = new UserSettings(null, null, null, settings, null, null, null)
            };
        }
    }
}
