using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("EntryLines")]
    public class EntryLine
    {
        [Key]
        public int RecID { get; set; }

        public int? LineNum { get; set; }

        public string ConsignmentID { get; set; }

        public long? DeclarationID { get; set; }

        public string LineID { get; set; }

        public string LineDesc { get; set; }

        public string LinePackType { get; set; }

        public string Location { get; set; }

        public int? LineQuantityDeclared { get; set; }

        public DateTime? LastUpdated { get; set; }

        public int? RequestID { get; set; }

        public string FormattedLineQuantityDeclared => LineQuantityDeclared?.ToString("N0") ?? string.Empty;

    }
}
