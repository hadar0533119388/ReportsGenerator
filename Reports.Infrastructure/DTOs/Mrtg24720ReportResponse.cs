using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class Mrtg24720ReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public List<Item> ItemsList { get; set; }

        public List<EntryLineMoveView> EntryLineMoveList { get; set; }

        public ReportDtl ReportDtl { get; set; }

        public int VarSequence { get; set; }

        public string SumWeight => ItemsList.Sum(item => item.ItemWeight)?.ToString("N0") ?? string.Empty;

        public int? SumQuantityMove => EntryLineMoveList.Sum(item => item.LineQuantityMove);

        public string FormattedSumQuantityMove => SumQuantityMove?.ToString("N0") ?? string.Empty;

        public double? UnitValue => Consignment.FobValueNis / ItemsList.Sum(item => item.Quantity);

        public string DepositValue => (UnitValue * SumQuantityMove)?.ToString("N0") ?? string.Empty;

        public string DateOpen => Consignment.OpeningDate?.ToString("dd/MM/yyyy HH:mm");



        public Mrtg24720ReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            ItemsList = new List<Item>();
            EntryLineMoveList = new List<EntryLineMoveView>();
        }
    }
}

