using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Repositories
{
    public class ReportRepositoryAdoNet : IReportRepositoryAdoNet
    {
        private readonly string connectionString;

        public ReportRepositoryAdoNet(string connectionString)
        {
            this.connectionString = connectionString;
        }

        public Task<DataSet> GetDataAsync(string reportID, Dictionary<string, object> Parameters)
        {
            throw new NotImplementedException();
        }
    }
}
