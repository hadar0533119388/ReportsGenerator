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
        Task<DataSet> GetDataAsync(string reportID, Dictionary<string, object> Parameters);
    }
}
