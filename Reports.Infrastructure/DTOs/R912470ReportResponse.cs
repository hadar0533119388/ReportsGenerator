using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R912470ReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public List<Item> ItemsList { get; set; }

        public List<Control1050> Control1050List { get; set; }

        public string TotalSum => ItemsList.Sum(item => item.Quantity)?.ToString("N0") ?? string.Empty;        

        public string ContainersNumber => string.Join(", ", Control1050List.Select(item => item.ContainerNumber));



        public R912470ReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            ItemsList = new List<Item>();
            Control1050List = new List<Control1050>();
        }
    }
}
