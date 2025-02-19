using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Exceptions;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.ReportGenerator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Core.Services
{
    public class ReportService : IReportService
    {
        private readonly IReportGeneratorFactory reportGeneratorFactory;
        private readonly ILogger logger;


        public ReportService(IReportGeneratorFactory reportGeneratorFactory, ILogger logger)
        {
            this.reportGeneratorFactory = reportGeneratorFactory;
            this.logger = logger;
        }

        public async Task<byte[]> GetReportAsync(ReportRequest request)
        {
            try
            {
                logger.WriteLog($"Generate Report: {request.ReportID}, ManifestID: {request.ManifestID}, PrinterName: {request.PrinterName}, User: {request.User}. - Process started");

                IReportGenerator reportGenerator = reportGeneratorFactory.GetReportGenerator(request.Type);
                byte[] reportBytes = await reportGenerator.ExecuteAsync(request);

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
