﻿using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("ConsignmentsReleasesItems")]
    public class ConsignmentReleaseItem
    {
        [Key]
        public int RecID { get; set; }

        public int? ConsignmentRecID { get; set; }

        public short? ItemSequence { get; set; }

        public string ManifestID { get; set; }

        public string ConsignmentID { get; set; }

        public string SourceConsignment { get; set; }

        public string CargoDescription { get; set; }

        public double? ItemWeight { get; set; }

        public int? Quantity { get; set; }

        public string WCOPackageType { get; set; }

        public string PackageType { get; set; }

        public string HScode { get; set; }

        public string AgentFileReferenceID { get; set; }

        public string MarksAndNumbers { get; set; }

        public bool? IsLastReleaseIndication { get; set; }

        public string FormattedItemWeight => ItemWeight?.ToString("N0") ?? string.Empty;

        public string FormattedQuantity => Quantity?.ToString("N0") ?? string.Empty;

    }
}
