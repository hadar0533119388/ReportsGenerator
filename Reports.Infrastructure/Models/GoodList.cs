using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("GoodsList")]
    public class GoodList
    {
        [Key]
        public int? GoodsCode { get; set; }

        public string GoodsName { get; set; }

        public string TaxGrp { get; set; }
    }
}
