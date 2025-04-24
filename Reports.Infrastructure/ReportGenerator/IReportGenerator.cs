using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.ReportGenerator
{
    public interface IReportGenerator
    {
        ReportType Type { get; }
        Task<byte[]> ExecuteAsync(ReportRequest request, ReportDtl reportDtl);
    }
}
