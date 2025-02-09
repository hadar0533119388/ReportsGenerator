using Reports.Infrastructure.DTOs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Core.Services
{
    public interface IReportService
    {
        Task<byte[]> GetReportAsync(ReportRequest request);
    }
}
