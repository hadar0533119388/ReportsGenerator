using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R60ExOutReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public List<EntryLineMoveView> EntryLineMoveList { get; set; }

        public string DateOpen => DateTime.Now.ToString("dd/MM/yyyy HH:mm");

        public ReportDtl ReportDtl { get; set; }

        public int VarSequence { get; set; }

        public SpecialAct SpecialAct { get; set; }

        public string DriverName { get; set; }

        public string DriverID { get; set; }

        public string LicensePlate { get; set; }

        public string SumExDelivered => EntryLineMoveList.Sum(item => item.exDelivered)?.ToString("N0") ?? string.Empty;



        public R60ExOutReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            EntryLineMoveList = new List<EntryLineMoveView>();
            ReportDtl = new ReportDtl();
            SpecialAct = new SpecialAct();

        }
    }
}
