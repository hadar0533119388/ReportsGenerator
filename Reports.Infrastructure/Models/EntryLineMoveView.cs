﻿using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("EntryLinesMovesView")]
    public class EntryLineMoveView
    {
        [Key]
        public int RecID { get; set; }

        public int? InvMovRecID { get; set; }

        public int? LineNum { get; set; }

        public string ConsignmentID { get; set; }

        public long? DeclarationID { get; set; }

        public int? LineQuantityMove { get; set; }

        public DateTime? LastUpdated { get; set; }

        public int? RequestIdOnMove { get; set; }

        public string LineID { get; set; }

        public string LineDesc { get; set; }

        public string LinePackType { get; set; }

        public string Location { get; set; }

        public int? LineQuantityDeclared { get; set; }

        public int? RequesIdforSplited { get; set; }

        public short? Mortgaged { get; set; }

        public string exPackageType { get; set; }

        public int? exReturned { get; set; }

        public int? exDelivered { get; set; }

        public int? Expr1 { get; set; }

        public string FormattedLineQuantityDeclared => LineQuantityDeclared?.ToString("N0") ?? string.Empty;

        public string FormattedLineQuantityMove => LineQuantityMove?.ToString("N0") ?? string.Empty;

        public int SerialNumber { get; set; }


    }
}
