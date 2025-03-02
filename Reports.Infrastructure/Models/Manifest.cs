using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("Manifests")]
    public class Manifest
    {
        [Key]
        public int RecID { get; set; }

        public string ManifestID { get; set; }

        public int? TermVAT { get; set; }

        public string TermUN { get; set; }

        public string GroupName { get; set; }

        public int? GroupID { get; set; }

        public string Location { get; set; }

        public int? ZIPcode { get; set; }

        public int? POB { get; set; }

        public string Phone { get; set; }

        public string ConsumerId { get; set; }

        public string TermName { get; set; }

        public string StylesheetName { get; set; }

        public string OutMessages { get; set; }

        public string InMessages { get; set; }

        public string MailNotifierAddress { get; set; }

        public string URLLink { get; set; }

        public string AttachmentsPath { get; set; }

        public int? LastGush { get; set; }

        public string OfficePrinter { get; set; }
    }
}
