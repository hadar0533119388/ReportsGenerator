using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("CustomersList")]

    public class CustomersList
    {
        public int? Code { get; set; }

        public int? BillingNo { get; set; }

        public long? VAT { get; set; }

        public string Name { get; set; }

        public short? Type { get; set; }

        public string BlockCode { get; set; }

    }

}
