using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json_Excel_IIS_Conversion_App.Models
{
    public class Sites
    {
        public string? SiteName { get; set; }
        public bool FatalErrorFound { get; set; }
        public FailedChecks[]? FailedChecks  { get; set; }
        public WarningChecks[]? WarningChecks { get; set; }
        public string? ManagedPipelineMode { get; set; }
        public bool Is32Bit { get; set; }
        public string? NetFrameworkVersion { get; set; }
        public VirtualApplications[]? VirtualApplications { get; set; }
    }
}
