using System.Collections.Generic;

namespace infotron.VbaCodeAnalizer.Inspections
{
    public class Issue
    {
        public string Type { get; set; }
        public string ModuleName { get; set; }
        public string Severity { get; set; }
        public string Description { get; set; }
        public string Name { get; set; }
        public string Meta { get; set; }
        public int Line { get; set; }
        public int Column { get; set; }
        public string FileName { get; set; }
    }

    public class CodeInspection
    {
        public List<Issue> MaintabilityAndReadabilityIssues { get; set; } = new List<Issue>();
    }

    public enum Helper
    {
        Yes,
        No
    }
}
