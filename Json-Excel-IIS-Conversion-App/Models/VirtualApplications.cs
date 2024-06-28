using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Json_Excel_IIS_Conversion_App.Models
{
    public class VirtualApplications
    {
        public string virtualPath { get; set; }
        public string physicalPath { get; set; }
        public virtualDirectories[] virtualDirectories { get; set; }

    }
}
