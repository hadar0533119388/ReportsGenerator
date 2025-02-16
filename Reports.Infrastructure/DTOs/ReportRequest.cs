using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.DTOs
{
    public class ReportRequest
    {
        [Required]
        public ReportType Type { get; set; }
        [Required]
        public string ReportID { get; set; }
        [Required]
        public string ManifestID { get; set; }
        [Required]
        public Dictionary<string, object> Parameters { get; set; }
        public bool IsPrint { get; set; }
        public string PrinterName { get; set; }
        [Required]
        public string User { get; set; }
        [Required]
        public string ConnectionString { get; set; }



        public ReportRequest()
        {
            Parameters = new Dictionary<string, object>();
        }
    }
}
