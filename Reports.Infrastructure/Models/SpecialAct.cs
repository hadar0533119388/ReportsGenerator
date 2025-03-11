using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("SpecialAct")]
    public class SpecialAct
    {
        [Key]
        public int RecID { get; set; }

        public string ManifestID { get; set; }

        public string ConsignmentID { get; set; }

        public int? RequestID { get; set; }

        public DateTime? DecisionDate { get; set; }

        public short? ActivityType { get; set; }

        public string ActivityTypeDesc { get; set; }

        public string AdditionalDesc { get; set; }

        public short? RequesterType { get; set; }

        public short? AuthorityCode { get; set; }

        public short? ActivityStat { get; set; }

        public DateTime? CommitTime { get; set; }

        public string PerformedBy { get; set; }

        public string Remark { get; set; }

        public short? SpecialActivityTypeEssence { get; set; }

        public string ResponseStat60 { get; set; }

        public string ResponseID60 { get; set; }

        public string ResponseErr60 { get; set; }

        public DateTime? TransmissionTime60 { get; set; }

        public DateTime? LastSuccessTrans60 { get; set; }

        public string AuthorityDesc { get; set; }

    }
}
