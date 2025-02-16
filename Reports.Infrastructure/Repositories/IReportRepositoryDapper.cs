using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Repositories
{
    public interface IReportRepositoryDapper
    {
        Task<dynamic> GetDataAsync(ReportRequest request, ReportDtl reportDtl);        
        Task<Manifest> GetManifestByManifestIDAsync(string manifestID);
        Task<ReportDtl> GetReportsDtlByReportIDAsync(string reportID);
        Task<R912470ReportResponse> GetDataForR912470Report(Dictionary<string, object> parameters, Manifest manifest);
        Task<R2470ReportResponse> GetDataForR2470Report(Dictionary<string, object> parameters, Manifest manifest);



    }
}
