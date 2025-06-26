using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("InventoryMove")]

    public class InventoryMove
    {
        [Key]
        public int RecID { get; set; }

        public int ItemRecID { get; set; }

        public short ItemSequence { get; set; }

        public string Consignment { get; set; }

        public string Manifest { get; set; }

        public long? DeclarationID { get; set; }

        public string MoveType { get; set; }

        public string ActionCode { get; set; }

        public int? AppearanceID { get; set; }

        public string TruckID { get; set; }

        public string TrailerID { get; set; }

        public short? CargoTransferMethodType { get; set; }

        public int? DriverID { get; set; }

        public string TransportationCompany { get; set; }

        public DateTime? MoveDate { get; set; }

        public double? WeightOnMove { get; set; }

        public int? QuantityOnMove { get; set; }

        public string ExceptionType { get; set; }

        public string ExceptionDetail { get; set; }

        public string ResponseStat { get; set; }

        public string ResponseID { get; set; }

        public string ResponseErr { get; set; }

        public DateTime? TransmissionTime { get; set; }

        public DateTime? LastSuccessTransmission { get; set; }

        public DateTime? CancelDate { get; set; }

        public string MovementFileName { get; set; }

        public string ReportingUser { get; set; }

        public int? Retries { get; set; }

        public DateTime? FirstTransmitionTime { get; set; }

        public int? RequestID { get; set; }

        public int? UnitedMovRef { get; set; }

        public string FormattedQuantityOnMove => QuantityOnMove?.ToString("N0") ?? string.Empty;

        public string FormattedMoveDate => MoveDate?.ToString("dd/MM/yy");



    }
}
