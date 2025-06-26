using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class InvRepForCustomsReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public Item Item { get; set; }

        public InventoryMove InventoryMoveEntry { get; set; }

        public InventoryMove InventoryMoveAdd { get; set; }

        public List<ConsignmentRelease> ConsignmentReleaseList { get; set; }

        public ReportDtl ReportDtl { get; set; }        

        public string DateOpen => DateTime.Now.ToString("dd/MM/yyyy HH:mm");


        public InvRepForCustomsReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            Item = new Item();
            InventoryMoveEntry = new InventoryMove();
            InventoryMoveAdd = new InventoryMove();
            ConsignmentReleaseList = new List<ConsignmentRelease>();
            ReportDtl = new ReportDtl();

        }

    }
}
