using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R60splitReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public List<EntryLine> EntryLineList { get; set; }

        public string DateOpen => DateTime.Now.ToString("dd/MM/yyyy HH:mm");

        public ReportDtl ReportDtl { get; set; }

        public int VarSequence { get; set; }

        public string SumLineQuantityDeclared => EntryLineList.Sum(item => item.LineQuantityDeclared)?.ToString("N0") ?? string.Empty;



        public R60splitReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            EntryLineList = new List<EntryLine>();
            ReportDtl = new ReportDtl();
        }
    }
}
