using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("ReportsDtl")]
    public class ReportDtl
    {
        [Key]
        public string ReportID { get; set; }

        public string ReportName { get; set; }

        public string Description { get; set; }

        public string Template { get; set; }

        public string FunctionName { get; set; }
    }
}
