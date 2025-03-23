using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Repositories
{
    public interface IReportRepositoryAdoNet
    {
        DataSet GetData(ReportRequest request, ReportDtl reportDtl);
    }
}
