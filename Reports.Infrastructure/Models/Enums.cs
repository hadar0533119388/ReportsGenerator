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
            GetDataForR1050MTReport,
            GetDataForR2470outReport,
            GetDataForMrtg24720Report,
            GetDataForR24720PReport,
            GetDataForR60ExOutReport,
            GetDataForR60ExInReport,
            GetDataForR60splitReport,
            GetDataForSUMentries9Report,
            GetDataForInvBckReport,
            GetDataForSUMvalindex3Report,
            GetDataForSUMdeliveryGush8Report,
            GetDataForSUMdeliveryLines8Report,
            GetDataForR2470outCollectReport,
            GetDataForCarsInShowroomsReport
        }

        public enum ReportID
        {
            SUMentries9,
        }

        public enum GenerateExcel
        {
            GenerateSUMentries9Report,
            GenerateInvBckReport,
            GenerateSUMvalindex3Report,
            GenerateSUMdeliveryGush8Report,
            GenerateSUMdeliveryLines8Report,
            GenerateCarsInShowroomsReport
        }
    }
}
