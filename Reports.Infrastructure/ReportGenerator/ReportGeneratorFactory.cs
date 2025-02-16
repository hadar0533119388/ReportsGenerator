using Reports.Infrastructure.Exceptions;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.ReportGenerator
{
    public class ReportGeneratorFactory : IReportGeneratorFactory
    {
        private readonly IDictionary<ReportType, IReportGenerator> reportGeneratorsMap;
        private readonly ILogger logger;

        public ReportGeneratorFactory(IEnumerable<IReportGenerator> reportGenerators, ILogger logger) 
        {
            //dependency injection, put in the disctionary all reportGenerators types (PDF/EXCEL)
            reportGeneratorsMap = reportGenerators.ToDictionary(report => report.Type, report => report);
            this.logger = logger;
        }

        public IReportGenerator GetReportGenerator(Enums.ReportType type)
        {
            try
            {
                reportGeneratorsMap.TryGetValue(type, out var reportGenerator);
                return reportGenerator;
            }
            catch(Exception ex)
            {
                logger.WriteLog($"Exception to get report generators: {ex.Message}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }
    }

}

/*
Dictionary<Key, value>

Dictionary<"PDF", ReportGeneratorInstance>
Dictionary<"EXCEL", ReportGeneratorInstance>

*/