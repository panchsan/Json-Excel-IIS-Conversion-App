using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json_Excel_IIS_Conversion_App.Models
{
    public class FailedChecks
    {
        public string IssueId { get; set; }
        public string detailsString { get; set; }
        public string Status { get; set; }
        public string Description { get; set; }
        public string Details { get; set; }
        public string Recommendation { get; set; }
        public string MoreInfoLink { get; set; }

    }
}
