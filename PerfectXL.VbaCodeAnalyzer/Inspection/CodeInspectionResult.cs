using System.Collections.Generic;

namespace PerfectXL.VbaCodeAnalyzer.Inspection
{
    public class CodeInspectionResult
    {
        public CodeInspectionResult(string moduleName)
        {
            ModuleName = moduleName;
        }

        public string ModuleName { get; }
        public List<VbaCodeIssue> VbaCodeIssues { get; set; } = new List<VbaCodeIssue>();
    }
}