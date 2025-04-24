using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Exceptions;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.ReportGenerator;
using Reports.Infrastructure.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Core.Services
{
    public class ReportService : IReportService
    {
        private readonly IReportGeneratorFactory reportGeneratorFactory;
        private readonly ILogger logger;
        private readonly IReportRepositoryDapper repositoryDapper;



        public ReportService(IReportGeneratorFactory reportGeneratorFactory, ILogger logger, IReportRepositoryDapper repositoryDapper)
        {
            this.reportGeneratorFactory = reportGeneratorFactory;
            this.logger = logger;
            this.repositoryDapper = repositoryDapper;
        }

        public async Task<byte[]> GetReportAsync(ReportRequest request)
        {
            try
            {
                string parameters = request.Parameters != null && request.Parameters.Any() ? string.Join(", ", request.Parameters.Select(kv => $"{kv.Key}: {kv.Value}")): "No parameters";

                logger.WriteLog($"Generate Report: {request.ReportID}, ManifestID: {request.ManifestID}, PrinterName: {request.PrinterName}, User: {request.User}, Parameters: {parameters}. - Process started");

                //Get global data from ReportsDtl table
                var reportDtl = await repositoryDapper.GetReportsDtlByReportIDAsync(request.ReportID);

                if (reportDtl == null)
                    throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);

                if (!Enum.TryParse<ReportType>(reportDtl.ReportFormat, out var reportType))
                {
                    throw new ArgumentException($"Invalid report format: {reportDtl.ReportFormat}");
                }

                IReportGenerator reportGenerator = reportGeneratorFactory.GetReportGenerator(reportType);
                byte[] reportBytes = await reportGenerator.ExecuteAsync(request, reportDtl);

                int status = reportBytes == null ? -1 : 0;
                logger.WriteLog($"Generate Report: {request.ReportID}, ManifestID: {request.ManifestID}, PrinterName: {request.PrinterName}, User: {request.User}. - Process completed with status: {status}");
                return reportBytes;
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Generate Report: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Generate Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }
    }
}
