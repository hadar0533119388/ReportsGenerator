using Autofac;
using Reports.Core.Services;
using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Exceptions;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.ReportGenerator;
using Reports.Infrastructure.Repositories;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Core.Configuration
{
    public static class ReportGeneratorFacade
    {        
        public static async Task<byte[]> GenerateReportAsync(ReportRequest request)
        {
            try
            {
                var builder = new ContainerBuilder();

                string logDirectory = ConfigurationManager.AppSettings["ReportsLogPath"];
                string sumatraPDF = ConfigurationManager.AppSettings["SumatraPDF"];
                string chrome = ConfigurationManager.AppSettings["Chrome"];
                string libreOffice = ConfigurationManager.AppSettings["LibreOffice"];

                builder.RegisterType<ReportRepositoryAdoNet>().As<IReportRepositoryAdoNet>()
                       .WithParameter("connectionString", request.ConnectionString)
                       .InstancePerDependency();

                builder.RegisterType<ReportRepositoryDapper>().As<IReportRepositoryDapper>()
                       .WithParameter("connectionString", request.ConnectionString)
                       .InstancePerDependency();

                builder.RegisterType<PDFReportGenerator>().As<IReportGenerator>()
                       .WithParameter("sumatraPDF", sumatraPDF)
                       .WithParameter("chrome", chrome)
                       .InstancePerDependency();

                builder.RegisterType<ExcelReportGenerator>().As<IReportGenerator>()
                       .WithParameter("libreOffice", libreOffice)
                       .InstancePerDependency();
                builder.RegisterType<ReportGeneratorFactory>().As<IReportGeneratorFactory>().SingleInstance();
                builder.RegisterType<ReportService>().As<IReportService>().InstancePerDependency();
                builder.RegisterType<Logger>().As<ILogger>().WithParameter("logDirectory", logDirectory).SingleInstance();

                using (var container = builder.Build())
                using (var scope = container.BeginLifetimeScope())
                {
                    var reportService = scope.Resolve<IReportService>();
                    return await reportService.GetReportAsync(request);
                }
            }
            catch (CustomException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }
    }
}
