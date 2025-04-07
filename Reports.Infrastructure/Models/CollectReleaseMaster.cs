using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("CollectReleaseMaster")]
    public class CollectReleaseMaster
    {
        public string ManifestID { get; set; }

        [Key]
        public int UnitedMovRef { get; set; }

        public string CustomerRef { get; set; }

        public int? ImporterID { get; set; }

        public string OrderDesc { get; set; }

        public DateTime? OrderDate { get; set; }

        public DateTime? DeliveryDate { get; set; }

        public int? DriverID { get; set; }

        public string DriverName { get; set; }

        public string TruckID { get; set; }

        public string Remarks4Delively { get; set; }

        public short? Stat { get; set; }

    }
}
