using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.ReportGenerator
{
    public interface IReportGeneratorFactory
    {
        IReportGenerator GetReportGenerator(ReportType type);
    }
}
