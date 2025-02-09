using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("ConsignmentsReleases")]
    public class ConsignmentRelease
    {
        [Key]
        public int RecID { get; set; }

        public int? ManifestRecID { get; set; }

        public string ManifestID { get; set; }

        public string ConsignmentID { get; set; }

        public string CargoIdentifierType { get; set; }

        public string CargoIdentifierKey1 { get; set; }

        public string CargoIdentifierKey2 { get; set; }

        public string CargoIdentifierKey3 { get; set; }

        public string ConsignmentStatus { get; set; }

        public DateTime? ReceiveDate { get; set; }

        public DateTime? LastUpdatedDate { get; set; }

        public string FatherConsignment { get; set; }

        public string DeliverySiteNumber { get; set; }

        public string ImporterID { get; set; }

        public long? DeclarationID { get; set; }

        public string ConsignmentDescription { get; set; }

        public short? ReleaseMessageCode { get; set; }

        public DateTime? ReleaseDate { get; set; }

        public string CustomsAgent { get; set; }

        public string CustomsAgentID { get; set; }

        public string CustomsAgentFile { get; set; }

        public string ImporterName { get; set; }

        public double? CifValueNis { get; set; }

        public double? DealValueNIS { get; set; }

        public double? TotalTax { get; set; }

        public double? FobValueNis { get; set; }

        public int? GovernmentProcedureType { get; set; }

        public int? DeclarationMessageType { get; set; }

        public int? ApplicationID { get; set; }

        public double? DeclarationVer { get; set; }

        public string PackageID { get; set; }

        public string FileName { get; set; }

        public string InternalRemarks { get; set; }

        public string CustomerRemarks { get; set; }



    }
}
