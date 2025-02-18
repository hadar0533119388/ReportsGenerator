using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R2470ReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public ConsignmentRelease ConsignmentRelease { get; set; }

        public List<ConsignmentReleaseItem> ConsignmentReleaseItemList { get; set; }

        public string TotalSum => ConsignmentReleaseItemList.Sum(item => item.Quantity)?.ToString("N0") ?? string.Empty;

        public string CargoDescription => string.Join(", ", ConsignmentReleaseItemList.Select(item => item.CargoDescription));

        public bool IsCustomsAgentIDNotSame => Consignment.CustomsAgentID != ConsignmentRelease.CustomsAgentID;

        public string DatePrint => DateTime.Now.ToString("dd/MM/yyyy HH:mm");



        public R2470ReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            ConsignmentRelease = new ConsignmentRelease();
            ConsignmentReleaseItemList = new List<ConsignmentReleaseItem>();
        }
    }
}
