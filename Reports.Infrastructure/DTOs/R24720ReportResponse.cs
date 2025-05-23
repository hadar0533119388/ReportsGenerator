﻿using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R24720ReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public List<Item> ItemsList { get; set; }

        public List<Control1050> Control1050List { get; set; }

        public List<EntryLineMoveView> EntryLineMoveList { get; set; }

        public string SumWeight => ItemsList.Sum(item => item.ItemWeight)?.ToString("N0") ?? string.Empty;

        public string SumQuantityMove => EntryLineMoveList.Sum(item => item.LineQuantityMove)?.ToString("N0") ?? string.Empty;

        public string DateOpen => Consignment.OpeningDate?.ToString("dd/MM/yyyy HH:mm");

        public ReportDtl ReportDtl { get; set; }

        public int VarSequence { get; set; }



        public R24720ReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            ItemsList = new List<Item>();
            Control1050List = new List<Control1050>();
            EntryLineMoveList = new List<EntryLineMoveView>();
        }
    }
}
