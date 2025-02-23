using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R1050MTReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public Control1050 Control1050 { get; set; }


        public R1050MTReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            Control1050 = new Control1050();
        }
    }
}
