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
        Task<R24720ReportResponse> GetDataForR24720Report(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl);
        Task<R1050MTReportResponse> GetDataForR1050MTReport(Dictionary<string, object> parameters, Manifest manifest);
        Task<R2470outReportResponse> GetDataForR2470outReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl, string User);
        Task<Mrtg24720ReportResponse> GetDataForMrtg24720Report(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl);
        Task<R24720PReportResponse> GetDataForR24720PReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl);
        Task<R60ExOutReportResponse> GetDataForR60ExOutReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl);


    }
}
