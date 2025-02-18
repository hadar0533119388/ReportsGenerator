using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("control1050")]

    public class Control1050
    {
        [Key]
        public int RecID { get; set; }

        public string Manifest { get; set; }

        public string ConsignmentID { get; set; }

        public string DocumentNumber { get; set; }

        public string ContainerNumber { get; set; }

        public string ExitedFrom { get; set; }

        public DateTime? OnTheWayTime { get; set; }

        public DateTime? MoveDate { get; set; }

        public int? SealCompletenessCode { get; set; }

        public string SealNumber { get; set; }

        public int? FunctionCode { get; set; }

        public int? DriverID { get; set; }

        public string TruckID { get; set; }

        public int? QuantityOnMove { get; set; }

        public string MsgContnt { get; set; }

        public string ResponseStat { get; set; }

        public string ResponseID { get; set; }

        public string ResponseErr { get; set; }

        public DateTime? TransmissionTime { get; set; }

        public DateTime? LastSuccessTransmission { get; set; }

        public DateTime? MTMoveDate { get; set; }

        public DateTime? MTingDate { get; set; }

        public short? MTingType { get; set; }

        public string Seal2Number { get; set; }

        public short? MTingTypeMessage { get; set; }

    }
}
