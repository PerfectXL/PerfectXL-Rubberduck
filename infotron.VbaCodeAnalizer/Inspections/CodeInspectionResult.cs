using System.Collections.Generic;

namespace infotron.VbaCodeAnalizer.Inspections
{
    public class CodeInspectionResult
    {
        public List<VbaCodeIssue> VbaCodeIssues { get; set; } = new List<VbaCodeIssue>();
    }
}