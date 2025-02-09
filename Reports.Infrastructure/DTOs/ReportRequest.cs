using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.DTOs
{
    public class ReportRequest
    {
        public ReportType Type { get; set; }
        public string ReportID { get; set; }
        public string ManifestID { get; set; }
        public Dictionary<string, object> Parameters { get; set; }
        public bool IsPrint { get; set; }
        public string PrinterName { get; set; }
        public string User { get; set; }
        public string ConnectionString { get; set; }



        public ReportRequest()
        {
            Parameters = new Dictionary<string, object>();
        }
    }
}
