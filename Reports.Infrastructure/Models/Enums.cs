using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    public class Enums
    {
        public enum ReportType
        {
            PDF = 1,
            Excel = 2
        }
        public enum StoredProcedure
        {
            GetManifestByManifestID,
            GetReportsDtl,
            GetDataForR912470Report,
            GetDataForR2470Report,
            GetDataForR24720Report,
            GetDataForR1050MTReport
        }
    }
}
